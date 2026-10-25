// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Threading.Channels;

namespace XlDuck.Core;

public sealed record SessionOptions(int MaxPending = 64, int HistoricalResults = 32, int MaxCaptureCells = 1000000, int MaxInputCharacters = 8 * 1024 * 1024);
public sealed record SessionStatistics(int Nodes, int ReferencedNodes, int Pending, int Historical, int Retiring, int Executions);

/// <summary>A lease belongs to a live formula/definition, independently of transient Excel subscriptions.</summary>
public sealed class ResultLease : IDisposable
{
    private WorkbookSession? _owner;
    public ResultHandle Handle { get; }
    public Task<ComputationResult> Completion { get; }
    internal ResultLease(WorkbookSession owner, ResultHandle handle, Task<ComputationResult> completion)
        => (_owner, Handle, Completion) = (owner, handle, completion);
    public Task<ComputationResult> WaitAsync(CancellationToken token = default) => Completion.WaitAsync(token);
    public void Dispose() => Interlocked.Exchange(ref _owner, null)?.Release(Handle);
}

public sealed class WorkbookSession : IAsyncDisposable
{
    private sealed class Entry(string key, ComputationRequest request, ResultHandle handle, PreparedQuery? prepared, Entry[] dependencies)
    {
        public readonly string Key = key;
        public readonly ComputationRequest Request = request;
        public readonly ResultHandle Handle = handle;
        public readonly PreparedQuery? Prepared = prepared;
        public Entry[] Dependencies = dependencies;
        public readonly TaskCompletionSource<ComputationResult> Completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly CancellationTokenSource Cancellation = new();
        public LinkedListNode<Entry>? Historical;
        public int References;
        public bool Stored;
        public bool Retired;
    }

    private readonly object _sync = new();
    private readonly SessionOptions _options;
    private readonly Lazy<Task<IExecutionBackend>> _backend;
    private readonly Dictionary<string, Entry> _requests = new(StringComparer.Ordinal);
    private readonly Dictionary<ResultHandle, Entry> _handles = new();
    private readonly LinkedList<Entry> _history = new();
    private readonly HashSet<Task> _cleanup = new();
    private readonly HashSet<Task<TablePage>> _reads = new();
    private readonly HashSet<CancellationTokenSource> _readCancellations = new();
    private TaskCompletionSource? _resume;
    private readonly Channel<Entry> _queue;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly Task _runner;
    private Task? _disposal;
    private bool _closed;
    private Exception? _cleanupError;
    private int _pending;
    private int _executions;
    public Guid Id { get; } = Guid.NewGuid();
    public bool IsClosed { get { lock (_sync) return _closed; } }
    public bool IsPaused { get { lock (_sync) return _resume is not null; } }

    public void SetPaused(bool paused)
    {
        lock (_sync)
        {
            EnsureOpen();
            if (paused) _resume ??= new(TaskCreationOptions.RunContinuationsAsynchronously);
            else { _resume?.TrySetResult(); _resume = null; }
        }
    }

    public void CancelPending()
    {
        lock (_sync)
        {
            EnsureOpen();
            foreach (var entry in _handles.Values.Where(entry => !entry.Completion.Task.IsCompleted).ToArray()) entry.Cancellation.Cancel();
            foreach (var cancellation in _readCancellations.ToArray()) cancellation.Cancel();
        }
    }

    public ComputationRequest GetDefinition(ResultHandle handle) { lock (_sync) return Find(handle).Request; }

    /// <summary>Permit new attempts on the next calculation, including in manual mode.
    /// Existing failed handles remain immutable until their owners release them.</summary>
    public void InvalidateFailuresForRetry()
    {
        lock (_sync)
        {
            EnsureOpen();
            foreach (var (key, entry) in _requests.ToArray())
                if (entry.Completion.Task.IsFaulted || entry.Completion.Task.IsCanceled) _requests.Remove(key);
        }
    }

    public WorkbookSession(Func<Guid, IExecutionBackend> backendFactory, SessionOptions? options = null)
    {
        _options = options ?? new();
        if (_options.MaxPending < 1 || _options.HistoricalResults < 0 || _options.MaxCaptureCells < 1 || _options.MaxInputCharacters < 1)
            throw new ArgumentOutOfRangeException(nameof(options));
        _backend = new(() => Task.Run(() => backendFactory(Id)));
        _queue = Channel.CreateBounded<Entry>(new BoundedChannelOptions(_options.MaxPending) { SingleReader = true });
        _runner = Task.Run(RunAsync);
    }

    public SessionStatistics Statistics
    {
        get { lock (_sync) return new(_handles.Count, _handles.Values.Count(e => e.References > 0), _pending + _reads.Count, _history.Count, _cleanup.Count, _executions); }
    }

    public ResultLease Acquire(ComputationRequest request, bool retryFailed = false)
    {
        request.Validate();
        if (request.Capture is { } capture && (long)capture.Headers.Length * capture.Rows.Length > _options.MaxCaptureCells)
            throw new ArgumentException($"Capture exceeds the {_options.MaxCaptureCells} cell limit.");
        long characters = request.Sql.Length + (request.PlotTemplate?.Length ?? 0);
        characters += request.Arguments.Sum(argument => (long)(argument.Value?.Text.Length ?? 0));
        if (request.Capture is { } data)
            characters += data.Headers.Sum(header => (long)header.Length) + data.Rows.Sum(row => row.Sum(value => (long)value.Text.Length));
        if (request.PlotOptions is { } options) characters += options.Sum(pair => (long)pair.Key.Length + pair.Value.Length);
        if (characters > _options.MaxInputCharacters) throw new ArgumentException("Computation inputs exceed this session's size limit. Use a smaller capture or prepare the data outside Excel.");
        var key = request.Identity();
        lock (_sync)
        {
            EnsureOpen();
            if (_requests.TryGetValue(key, out var cached) && !(retryFailed && (cached.Completion.Task.IsFaulted || cached.Completion.Task.IsCanceled))) return Lease(cached);
            if (_pending + _cleanup.Count >= _options.MaxPending)
                throw new InvalidOperationException("This workbook's computation queue is full. Wait for pending work, then retry.");

            // Validate and resolve everything before acquiring any dependency references.
            var dependencies = request.Arguments.Where(a => a.Handle is not null).Select(a => Find(a.Handle!.Value)).Distinct().ToArray();
            PreparedQuery? prepared = null;
            if (request.Kind is RequestKind.Query or RequestKind.Fragment)
                prepared = SqlTemplate.Prepare(request, handle => Find(handle).Request, _options.MaxInputCharacters);
            ResultLease? plotData = null;
            try
            {
                if (request.Kind == RequestKind.Plot)
                {
                    plotData = AcquireData(request.Arguments[0].Handle!.Value);
                    dependencies = [Find(plotData.Handle)];
                }
                if (_pending + _cleanup.Count >= _options.MaxPending)
                    throw new InvalidOperationException("This workbook's computation queue is full. Wait for pending work, then retry.");
                var kind = request.Kind switch { RequestKind.Fragment => ResultKind.Fragment, RequestKind.Plot => ResultKind.Plot, _ => ResultKind.Table };
                var entry = new Entry(key, request, new(Id, Guid.NewGuid(), kind), prepared, dependencies);
                foreach (var dependency in dependencies) Retain(dependency);
                _requests[key] = entry;
                _handles.Add(entry.Handle, entry);
                _pending++;
                var lease = Lease(entry);
                if (!_queue.Writer.TryWrite(entry))
                {
                    _pending--;
                    entry.Completion.TrySetCanceled();
                    lease.Dispose();
                    throw new InvalidOperationException("Session queue closed unexpectedly.");
                }
                return lease;
            }
            finally { plotData?.Dispose(); }
        }
    }

    public ResultLease AcquireExisting(ResultHandle handle)
    {
        lock (_sync) { EnsureOpen(); return Lease(Find(handle)); }
    }

    /// <summary>Fragments become retained materializations. Plots own their materialized data.</summary>
    public ResultLease AcquireData(ResultHandle handle)
    {
        lock (_sync)
        {
            EnsureOpen();
            var entry = Find(handle);
            return handle.Kind switch
            {
                ResultKind.Fragment => Acquire(ComputationRequest.Query("SELECT * FROM ?", QueryArgument.Reference(handle))),
                ResultKind.Plot => Lease(entry.Dependencies[0]),
                _ => Lease(entry)
            };
        }
    }

    public Task<TablePage> ReadAsync(ResultHandle handle, ReadRequest request, CancellationToken token = default)
    {
        lock (_sync)
        {
            EnsureOpen();
            if (_reads.Count >= _options.MaxPending) throw new InvalidOperationException("This workbook's output queue is full. Retry after pending output finishes.");
            var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
            _readCancellations.Add(cancellation);
            var task = ReadCoreAsync(handle, request, cancellation.Token);
            _reads.Add(task);
            _ = task.ContinueWith(completed =>
            {
                lock (_sync) { _reads.Remove(completed); _readCancellations.Remove(cancellation); cancellation.Dispose(); }
            }, TaskScheduler.Default);
            return task;
        }
    }
    private async Task<TablePage> ReadCoreAsync(ResultHandle handle, ReadRequest request, CancellationToken token)
    {
        if (request.MaxRows < 1 || request.MaxCells < 1) throw new ArgumentOutOfRangeException(nameof(request));
        using var retained = AcquireExisting(handle);
        if (handle.Kind != ResultKind.Table) throw new ArgumentException("Acquire a materialized data lease before reading output.");
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(token, _lifetime.Token);
        await retained.Completion.WaitAsync(linked.Token).ConfigureAwait(false);
        var backend = await _backend.Value.ConfigureAwait(false);
        var page = await backend.ReadAsync(handle, request, linked.Token).ConfigureAwait(false);
        linked.Token.ThrowIfCancellationRequested();
        return page;
    }

    /// <summary>Cell tokens carry a unique result ID. Only this session may resolve them.</summary>
    public ResultHandle ResolveCellHandle(string text)
    {
        lock (_sync)
        {
            EnsureOpen();
            var handle = ResultHandle.FromCellText(text, Id);
            if (!_handles.TryGetValue(handle, out var entry) || entry.Retired)
                throw new ArgumentException("This handle is not available in this workbook session. It may belong to a different workbook or an earlier session. Recompute its source here.");
            return entry.Handle;
        }
    }

    private Entry Find(ResultHandle handle)
    {
        if (handle.Session != Id) throw new ArgumentException("This handle belongs to a different workbook session. Recompute its source in this workbook.");
        if (!_handles.TryGetValue(handle, out var entry) || entry.Retired) throw new ArgumentException("This result is no longer retained. Recompute its source.");
        return entry;
    }
    private void EnsureOpen()
    {
        if (_closed) throw new ObjectDisposedException(nameof(WorkbookSession), "This workbook session has closed.");
        if (_cleanupError is not null) throw new InvalidOperationException("Result cleanup failed. Restart this workbook's XlDuck session.", _cleanupError);
    }
    private ResultLease Lease(Entry entry) { Retain(entry); return new(this, entry.Handle, entry.Completion.Task); }
    private void Retain(Entry entry)
    {
        if (entry.Historical is not null) { _history.Remove(entry.Historical); entry.Historical = null; }
        entry.References++;
    }
    internal void Release(ResultHandle handle)
    {
        lock (_sync)
        {
            if (_closed || !_handles.TryGetValue(handle, out var entry)) return;
            Release(entry);
        }
    }
    private void Release(Entry entry)
    {
        if (--entry.References < 0) throw new InvalidOperationException("Result reference count underflow.");
        if (entry.References != 0) return;
        if (!entry.Completion.Task.IsCompleted)
        {
            if (_requests.GetValueOrDefault(entry.Key) == entry) _requests.Remove(entry.Key);
            entry.Cancellation.Cancel();
        }
        else RememberOrRetire(entry);
    }
    private void RememberOrRetire(Entry entry)
    {
        if (entry.References != 0 || entry.Retired) return;
        // Caching an unowned fragment/plot could retain an arbitrarily large dependency graph.
        // Only independent materialized results (and bounded failure records) enter history.
        if (entry.Handle.Kind != ResultKind.Table || entry.Completion.Task.IsCanceled || _options.HistoricalResults == 0) { Retire(entry); return; }
        entry.Historical ??= _history.AddLast(entry);
        while (_history.Count > _options.HistoricalResults) Retire(_history.First!.Value);
    }
    private void Retire(Entry entry)
    {
        entry.Retired = true;
        _handles.Remove(entry.Handle);
        if (_requests.GetValueOrDefault(entry.Key) == entry) _requests.Remove(entry.Key);
        if (entry.Historical is not null) _history.Remove(entry.Historical);
        entry.Historical = null;
        foreach (var dependency in entry.Dependencies) Release(dependency);
        entry.Dependencies = [];
        entry.Cancellation.Dispose();
        if (entry.Stored)
        {
            var cleanup = Task.Run(async () => await (await _backend.Value.ConfigureAwait(false)).DropAsync(entry.Handle).ConfigureAwait(false));
            _cleanup.Add(cleanup);
            _ = cleanup.ContinueWith(task =>
            {
                lock (_sync)
                {
                    if (task.IsFaulted) _cleanupError ??= task.Exception!.GetBaseException();
                    _cleanup.Remove(task);
                }
            }, TaskScheduler.Default);
        }
    }

    private async Task RunAsync()
    {
        await foreach (var entry in _queue.Reader.ReadAllAsync().ConfigureAwait(false))
        {
            using var linked = CancellationTokenSource.CreateLinkedTokenSource(entry.Cancellation.Token, _lifetime.Token);
            ComputationResult? result = null;
            Exception? failure = null;
            bool cancelled = false;
            try
            {
                linked.Token.ThrowIfCancellationRequested();
                Task resumed;
                lock (_sync) resumed = _resume?.Task ?? Task.CompletedTask;
                await resumed.WaitAsync(linked.Token).ConfigureAwait(false);
                foreach (var dependency in entry.Dependencies) await dependency.Completion.Task.WaitAsync(linked.Token).ConfigureAwait(false);
                if (entry.Request.Kind == RequestKind.Fragment)
                {
                    await (await _backend.Value.ConfigureAwait(false)).ValidateAsync(entry.Prepared!, linked.Token).ConfigureAwait(false);
                    result = new(entry.Handle, [], 0);
                }
                else if (entry.Request.Kind == RequestKind.Plot)
                {
                    var source = await entry.Dependencies[0].Completion.Task.ConfigureAwait(false);
                    result = source with { Handle = entry.Handle };
                }
                else
                {
                    var backend = await _backend.Value.ConfigureAwait(false);
                    Interlocked.Increment(ref _executions);
                    result = await backend.ExecuteAsync(new(entry.Handle, entry.Prepared, entry.Request.Capture), linked.Token).ConfigureAwait(false);
                    entry.Stored = true;
                    if (result.Handle != entry.Handle) throw new InvalidOperationException("Engine returned an unexpected result identity.");
                }
                linked.Token.ThrowIfCancellationRequested();
            }
            catch (OperationCanceledException) { cancelled = true; }
            catch (Exception error) { failure = error; }
            finally
            {
                lock (_sync)
                {
                    _pending--;
                    if (entry.Request.Kind is RequestKind.Query or RequestKind.Capture)
                    {
                        foreach (var dependency in entry.Dependencies) Release(dependency);
                        entry.Dependencies = [];
                    }
                    // Publish only after dependency cleanup, under the ownership lock. Disposal cannot
                    // race a successful completion into publishing a result with no remaining owner.
                    if (cancelled || linked.IsCancellationRequested) entry.Completion.TrySetCanceled();
                    else if (failure is not null) entry.Completion.TrySetException(failure);
                    else entry.Completion.TrySetResult(result!);
                    if (!_closed) RememberOrRetire(entry);
                }
            }
        }
    }

    public async Task WaitForIdleAsync()
    {
        while (true)
        {
            Task[] pending;
            lock (_sync) pending = [.. _handles.Values.Select(e => e.Completion.Task).Where(t => !t.IsCompleted), .. _cleanup, .. _reads];
            if (pending.Length == 0) return;
            try { await Task.WhenAll(pending).ConfigureAwait(false); } catch { /* Results carry their individual errors. */ }
            await Task.Yield();
        }
    }
    public ValueTask DisposeAsync()
    {
        lock (_sync)
        {
            _closed = true;
            _queue.Writer.TryComplete();
            _disposal ??= Task.Run(DisposeCoreAsync);
            return new(_disposal);
        }
    }
    private async Task DisposeCoreAsync()
    {
        _lifetime.Cancel();
        await _runner.ConfigureAwait(false);
        Task[] cleanup;
        lock (_sync) cleanup = [.. _cleanup, .. _reads];
        try { await Task.WhenAll(cleanup).ConfigureAwait(false); }
        catch { /* Disposing the entire backend below also releases tables whose individual drops failed. */ }
        finally
        {
            IExecutionBackend? backend = null;
            if (_backend.IsValueCreated)
                try { backend = await _backend.Value.ConfigureAwait(false); } catch { /* Startup failure was reported on the request. */ }
            try { if (backend is not null) await backend.DisposeAsync().ConfigureAwait(false); }
            finally
            {
                lock (_sync)
                {
                    foreach (var entry in _handles.Values) entry.Cancellation.Dispose();
                    _handles.Clear(); _requests.Clear(); _history.Clear(); _cleanup.Clear(); _reads.Clear();
                }
                _lifetime.Dispose();
            }
        }
    }
}
