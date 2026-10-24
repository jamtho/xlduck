// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;

namespace XlDuck.ExecutionPrototype;

public sealed class Session(Func<IBackend> backendFactory) : IAsyncDisposable
{
    private sealed class Entry(CancellationTokenSource cancellation, Task<Result> task)
    {
        public CancellationTokenSource Cancellation { get; } = cancellation;
        public Task<Result> Task { get; } = task;
        public int Waiters;
        private bool _disposed;
        public void Cancel() { lock (this) { if (!_disposed) Cancellation.Cancel(); } }
        public void Dispose() { lock (this) { if (!_disposed) { Cancellation.Dispose(); _disposed = true; } } }
    }
    private readonly Lazy<Task<IBackend>> _backend = new(() => Task.Run(backendFactory));
    private readonly object _sync = new();
    private readonly Dictionary<string, Entry> _cache = new();
    private readonly CancellationTokenSource _lifetime = new();
    private readonly SemaphoreSlim _capacity = new(64);
    private Task? _disposal;
    private bool _closed;
    private int _executions;
    public string Id { get; } = Guid.NewGuid().ToString("N");
    public int Executions => Volatile.Read(ref _executions);
    public bool IsClosed { get { lock (_sync) return _closed; } }
    public int Pending { get { lock (_sync) return _cache.Values.Count(entry => !entry.Task.IsCompleted); } }
    public int? WorkerProcessId => _backend.IsValueCreated && _backend.Value.IsCompletedSuccessfully &&
        _backend.Value.Result is WorkerBackend worker ? worker.ProcessId : null;
    public void CrashWorker()
    {
        if (_backend.IsValueCreated && _backend.Value.IsCompletedSuccessfully && _backend.Value.Result is WorkerBackend worker)
            worker.SimulateCrash();
        else throw new InvalidOperationException("No worker is running.");
    }
    private static string Key(Request request) => JsonSerializer.Serialize(request with { Id = Guid.Empty });

    public bool TryGetCompleted(Request request, out Result? result)
    {
        lock (_sync)
        {
            if (!_closed && _cache.TryGetValue(Key(request), out var entry))
            {
                if (entry.Task.IsCompletedSuccessfully) { result = entry.Task.Result; return true; }
                if (entry.Task.IsFaulted)
                {
                    result = new(Guid.Empty, Error: entry.Task.Exception!.GetBaseException().Message);
                    return true;
                }
            }
            result = null;
            return false;
        }
    }

    public Task<Result> GetAsync(Request request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var key = Key(request);
        Entry entry;
        lock (_sync)
        {
            if (_closed) throw new OperationCanceledException("Session has closed.");
            if (!_cache.TryGetValue(key, out entry!))
            {
                var cancellation = new CancellationTokenSource();
                entry = new(cancellation, ExecuteAsync(request, cancellation.Token));
                _cache.Add(key, entry);
            }
            entry.Waiters++;
        }
        return WaitAsync(key, entry, cancellationToken);
    }

    private async Task<Result> WaitAsync(string key, Entry entry, CancellationToken token)
    {
        try { return await entry.Task.WaitAsync(token).ConfigureAwait(false); }
        finally
        {
            bool abandon;
            lock (_sync)
            {
                abandon = --entry.Waiters == 0 && !entry.Task.IsCompleted;
                if ((abandon || entry.Task.IsCanceled) && _cache.GetValueOrDefault(key) == entry)
                    _cache.Remove(key);
            }
            // Removing one cell must not cancel a computation another live cell still needs.
            if (abandon)
            {
                entry.Cancel();
                _ = entry.Task.ContinueWith(_ => entry.Dispose(), TaskScheduler.Default);
            }
        }
    }

    private async Task<Result> ExecuteAsync(Request request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetime.Token);
        await _capacity.WaitAsync(linked.Token).ConfigureAwait(false);
        try
        {
            var backend = await _backend.Value.ConfigureAwait(false);
            Interlocked.Increment(ref _executions);
            var result = await backend.ExecuteAsync(request, linked.Token).ConfigureAwait(false);
            linked.Token.ThrowIfCancellationRequested();
            return result;
        }
        finally { _capacity.Release(); }
    }
    public ValueTask DisposeAsync()
    {
        lock (_sync)
        {
            _closed = true;
            _disposal ??= Task.Run(DisposeCoreAsync);
            return new(_disposal);
        }
    }
    private async Task DisposeCoreAsync()
    {
        _lifetime.Cancel();
        Entry[] entries;
        lock (_sync) { entries = _cache.Values.ToArray(); _cache.Clear(); }
        if (_backend.IsValueCreated) await (await _backend.Value.ConfigureAwait(false)).DisposeAsync().ConfigureAwait(false);
        foreach (var entry in entries)
        {
            try { await entry.Task.ConfigureAwait(false); } catch { /* Reported to its subscribers. */ }
            entry.Dispose();
        }
    }
}
