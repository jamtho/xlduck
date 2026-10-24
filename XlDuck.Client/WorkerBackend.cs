// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Concurrent;
using System.Diagnostics;
using XlDuck.Core;

namespace XlDuck.Client;

public sealed class WorkerBackend : IExecutionBackend
{
    private readonly Process _process;
    private readonly SessionDirectory _storage = new();
    private readonly ConcurrentDictionary<Guid, TaskCompletionSource<WorkerResponse>> _pending = new();
    private readonly TaskCompletionSource _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly SemaphoreSlim _writes = new(1);
    private readonly SemaphoreSlim _admission = new(64);
    private readonly Task _reader;
    private readonly object _disposeSync = new();
    private Task? _disposal;
    private Exception? _failure;
    public int ProcessId { get; }
    public string StoragePath => _storage.Path;

    public WorkerBackend(Guid sessionId, string workerPath, ConnectionSettings? settings = null)
    {
        try
        {
            (settings ??= new()).Validate();
            var start = new ProcessStartInfo(Path.GetFullPath(workerPath))
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
            };
            start.ArgumentList.Add("--session"); start.ArgumentList.Add(sessionId.ToString("N"));
            start.ArgumentList.Add("--spill"); start.ArgumentList.Add(_storage.Path);
            _process = Process.Start(start) ?? throw new IOException("Could not start XlDuck worker.");
        }
        catch { _storage.Dispose(); throw; }
        ProcessId = _process.Id;
        // Drain diagnostics without logging potentially sensitive SQL or values.
        _process.ErrorDataReceived += (_, _) => { };
        _process.BeginErrorReadLine();
        _reader = StartAsync(settings);
    }
    private async Task StartAsync(ConnectionSettings settings)
    {
        try
        {
            await _writes.WaitAsync().ConfigureAwait(false);
            try { await WorkerProtocol.WriteAsync(_process.StandardInput.BaseStream, new WorkerStartup(settings)).ConfigureAwait(false); }
            finally { _writes.Release(); }
        }
        catch (Exception error) { _ready.TrySetException(error); StopOwnedWorker(); }
        await ReadResponsesAsync().ConfigureAwait(false);
    }
    public async Task ValidateAsync(PreparedQuery query, CancellationToken token) =>
        _ = await CallAsync(new(Guid.NewGuid(), WorkerOperation.Validate, Query: query), token).ConfigureAwait(false);
    public async Task<ComputationResult> ExecuteAsync(BackendRequest request, CancellationToken token) =>
        (await CallAsync(new(Guid.NewGuid(), WorkerOperation.Execute, Execution: request), token).ConfigureAwait(false)).Result
        ?? throw new IOException("Worker response omitted the result.");
    public async Task<TablePage> ReadAsync(ResultHandle handle, ReadRequest request, CancellationToken token) =>
        (await CallAsync(new(Guid.NewGuid(), WorkerOperation.Read, Handle: handle, Read: request), token).ConfigureAwait(false)).Page
        ?? throw new IOException("Worker response omitted output data.");
    public async Task DropAsync(ResultHandle handle) =>
        _ = await CallAsync(new(Guid.NewGuid(), WorkerOperation.Drop, Handle: handle), CancellationToken.None).ConfigureAwait(false);
    public async Task<int> GetStoredTableCountAsync() =>
        (await CallAsync(new(Guid.NewGuid(), WorkerOperation.Inspect), CancellationToken.None).ConfigureAwait(false)).StoredTables
        ?? throw new IOException("Worker response omitted storage diagnostics.");

    private async Task<WorkerResponse> CallAsync(WorkerRequest request, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!_admission.Wait(0)) throw new InvalidOperationException("Worker request queue is full. Retry after pending work finishes.");
        var completion = new TaskCompletionSource<WorkerResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        try
        {
            await _ready.Task.WaitAsync(TimeSpan.FromSeconds(10), token).ConfigureAwait(false);
            if (!_pending.TryAdd(request.Id, completion)) throw new InvalidOperationException("Duplicate worker request ID.");
            if (Volatile.Read(ref _failure) is { } failure) throw new IOException("XlDuck worker is unavailable. Restart this workbook's XlDuck session.", failure);
            token.ThrowIfCancellationRequested();
            var sending = SendAsync(request);
            using var registration = token.Register(() => _ = CancelAsync(request.Id, completion.Task));
            await sending.ConfigureAwait(false);
            // Wait for the engine's cancellation acknowledgement before releasing source tables.
            var response = await completion.Task.ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            if (response.Cancelled) throw new OperationCanceledException("The query was cancelled.", token);
            if (response.Error is not null) throw new InvalidOperationException(response.Error);
            return response;
        }
        finally { _pending.TryRemove(request.Id, out _); _admission.Release(); }
    }
    private async Task SendAsync(WorkerRequest request)
    {
        await _writes.WaitAsync().ConfigureAwait(false);
        try { await WorkerProtocol.WriteAsync(_process.StandardInput.BaseStream, request).ConfigureAwait(false); }
        finally { _writes.Release(); }
    }
    private async Task CancelAsync(Guid id, Task completion)
    {
        if (completion.IsCompleted) return;
        var sending = SendAsync(new(id, WorkerOperation.Cancel));
        try
        {
            // The deadline covers transport too: a wedged worker may never read its cancel frame.
            if (await Task.WhenAny(completion, Task.Delay(TimeSpan.FromSeconds(5))).ConfigureAwait(false) != completion)
                StopOwnedWorker();
        }
        catch { /* The read loop reports a disconnected worker to all outstanding calls. */ }
        finally { try { await sending.ConfigureAwait(false); } catch { /* Broken or closed pipe. */ } }
    }
    private async Task ReadResponsesAsync()
    {
        Exception failure = new IOException("XlDuck worker exited. Restart this workbook's XlDuck session.");
        try
        {
            while (await WorkerProtocol.ReadAsync<WorkerResponse>(_process.StandardOutput.BaseStream).ConfigureAwait(false) is { } response)
            {
                if (response.Version != WorkerProtocol.Version) throw new IOException("XlDuck host and worker versions do not match.");
                if (response.Ready) { _ready.TrySetResult(); continue; }
                if (_pending.TryRemove(response.Id, out var completion)) completion.TrySetResult(response);
            }
        }
        catch (Exception error) { failure = new IOException("XlDuck worker communication failed. Restart this workbook's XlDuck session.", error); }
        finally
        {
            Volatile.Write(ref _failure, failure);
            _ready.TrySetException(failure);
            foreach (var item in _pending)
                if (_pending.TryRemove(item.Key, out var completion)) completion.TrySetException(failure);
        }
    }
    private void StopOwnedWorker()
    {
        try { if (!_process.HasExited) _process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { /* Already exited. */ }
    }
    public ValueTask DisposeAsync()
    {
        lock (_disposeSync) return new(_disposal ??= DisposeCoreAsync());
    }
    private async Task DisposeCoreAsync()
    {
        Volatile.Write(ref _failure, new ObjectDisposedException(nameof(WorkerBackend)));
        var closingInput = CloseInputAsync();
        try
        {
            // Acquiring the write lock must also be bounded when a large frame is stuck in the pipe.
            try { await Task.WhenAll(closingInput, _process.WaitForExitAsync()).WaitAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false); }
            catch (Exception error) when (error is TimeoutException or IOException or ObjectDisposedException)
            {
                StopOwnedWorker();
                await _process.WaitForExitAsync().ConfigureAwait(false);
            }
            try { await closingInput.ConfigureAwait(false); } catch (IOException) { }
            await _reader.ConfigureAwait(false);
        }
        finally { _process.Dispose(); _storage.Dispose(); }
    }
    private async Task CloseInputAsync()
    {
        await _writes.WaitAsync().ConfigureAwait(false);
        try { _process.StandardInput.Close(); }
        finally { _writes.Release(); }
    }
}
