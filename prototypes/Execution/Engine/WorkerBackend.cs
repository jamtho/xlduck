// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.Json;

namespace XlDuck.ExecutionPrototype;

public sealed class WorkerBackend : IBackend
{
    private readonly Process _process;
    private readonly ConcurrentDictionary<Guid, TaskCompletionSource<Result>> _pending = new();
    private readonly SemaphoreSlim _writeGate = new(1);
    private readonly Task _reader;
    private Exception? _failure;
    public int ProcessId => _process.Id;
    public void SimulateCrash() => _process.Kill(entireProcessTree: true);

    public WorkerBackend(string workerPath)
    {
        _process = Process.Start(new ProcessStartInfo(workerPath)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true
        }) ?? throw new InvalidOperationException("Could not start worker.");
        // Drain stderr to prevent a full pipe blocking the child. It is not a protocol stream.
        _process.ErrorDataReceived += (_, _) => { };
        _process.BeginErrorReadLine();
        _reader = ReadAsync();
    }
    private async Task SendAsync(Request request)
    {
        await _writeGate.WaitAsync().ConfigureAwait(false);
        try
        {
            await _process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(request)).ConfigureAwait(false);
            await _process.StandardInput.FlushAsync().ConfigureAwait(false);
        }
        finally { _writeGate.Release(); }
    }
    public async Task<Result> ExecuteAsync(Request request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var completion = new TaskCompletionSource<Result>(TaskCreationOptions.RunContinuationsAsynchronously);
        if (!_pending.TryAdd(request.Id, completion)) throw new InvalidOperationException("Duplicate request ID.");
        try
        {
            // Pair this check with publishing failure before the reader drains pending requests.
            if (Volatile.Read(ref _failure) is { } failure) throw new IOException("Worker is unavailable.", failure);
            await SendAsync(request).ConfigureAwait(false);
            using var registration = cancellationToken.Register(() =>
            {
                _ = CancelAsync(request.Id);
                completion.TrySetCanceled(cancellationToken);
            });
            return await completion.Task.ConfigureAwait(false);
        }
        finally { _pending.TryRemove(request.Id, out _); }
    }
    private async Task CancelAsync(Guid id)
    {
        try { await SendAsync(new(id, "cancel")).ConfigureAwait(false); }
        catch { /* Worker exit also fails all outstanding requests. */ }
    }
    private async Task ReadAsync()
    {
        Exception failure = new IOException("Worker exited before replying.");
        try
        {
            while (await _process.StandardOutput.ReadLineAsync().ConfigureAwait(false) is { } line)
            {
                var result = JsonSerializer.Deserialize<Result>(line) ?? throw new IOException("Invalid worker response.");
                if (_pending.TryRemove(result.Id, out var completion))
                {
                    if (result.Error is null) completion.TrySetResult(result);
                    else completion.TrySetException(new InvalidOperationException(result.Error));
                }
            }
        }
        catch (Exception ex) { failure = ex; }
        finally
        {
            Volatile.Write(ref _failure, failure);
            foreach (var item in _pending)
                if (_pending.TryRemove(item.Key, out var completion)) completion.TrySetException(failure);
        }
    }
    public async ValueTask DisposeAsync()
    {
        await _writeGate.WaitAsync().ConfigureAwait(false);
        try
        {
            _process.StandardInput.Close();
            if (!_process.HasExited) _process.Kill(entireProcessTree: true);
        }
        finally { _writeGate.Release(); }
        await _reader.ConfigureAwait(false);
        await _process.WaitForExitAsync().ConfigureAwait(false);
        _process.Dispose();
    }
}
