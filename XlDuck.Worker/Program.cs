// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Concurrent;
using XlDuck.Core;
using XlDuck.Engine;

if (args.Length != 4 || args[0] != "--session" || args[2] != "--spill" || !Guid.TryParseExact(args[1], "N", out var session) || !Directory.Exists(args[3])) return 2;
using var outputGate = new SemaphoreSlim(1);
using var input = Console.OpenStandardInput();
using var output = Console.OpenStandardOutput();
var startup = await WorkerProtocol.ReadAsync<WorkerStartup>(input) ?? throw new IOException("Missing startup settings.");
if (startup.Version != WorkerProtocol.Version) throw new IOException("Unsupported startup protocol version.");
startup.Settings.Validate();
await using var backend = new DuckDbBackend(session, args[3], startup.Settings);
var requests = new ConcurrentDictionary<Guid, CancellationTokenSource>();
var running = new ConcurrentDictionary<Guid, Task>();
await Send(new(Guid.Empty, Ready: true));
try
{
    while (await WorkerProtocol.ReadAsync<WorkerRequest>(input) is { } request)
    {
        if (request.Version != WorkerProtocol.Version) throw new IOException("Unsupported host protocol version.");
        if (request.Operation == WorkerOperation.Cancel)
        {
            if (requests.TryGetValue(request.Id, out var pending))
                lock (pending) { if (requests.ContainsKey(request.Id)) pending.Cancel(); }
            continue;
        }
        if (requests.Count >= 64) { await Send(new(request.Id, Error: "Worker request queue is full. Retry after pending work finishes.")); continue; }
        var cancellation = new CancellationTokenSource();
        if (!requests.TryAdd(request.Id, cancellation)) { cancellation.Dispose(); throw new IOException("Duplicate request ID."); }
        var task = Respond(request, cancellation);
        running[request.Id] = task;
        _ = task.ContinueWith(completed => { _ = completed.Exception; running.TryRemove(request.Id, out _); }, TaskScheduler.Default);
    }
}
finally
{
    foreach (var pair in requests)
        lock (pair.Value) { if (requests.ContainsKey(pair.Key)) pair.Value.Cancel(); }
    await Task.WhenAll(running.Values);
}
return 0;

async Task Send(WorkerResponse response)
{
    await outputGate.WaitAsync();
    try { await WorkerProtocol.WriteAsync(output, response); }
    finally { outputGate.Release(); }
}
async Task Respond(WorkerRequest request, CancellationTokenSource cancellation)
{
    WorkerResponse response;
    try
    {
        response = request.Operation switch
        {
            WorkerOperation.Execute => new(request.Id, Result: await backend.ExecuteAsync(request.Execution ?? throw new ArgumentException("Missing execution request."), cancellation.Token)),
            WorkerOperation.Read => new(request.Id, Page: await backend.ReadAsync(request.Handle ?? throw new ArgumentException("Missing handle."), request.Read ?? throw new ArgumentException("Missing read request."), cancellation.Token)),
            WorkerOperation.Drop => await Drop(request),
            WorkerOperation.Inspect => new(request.Id, StoredTables: backend.TableCount),
            WorkerOperation.Validate => await Validate(request, cancellation.Token),
            _ => throw new ArgumentException("Unsupported worker operation.")
        };
    }
    catch (OperationCanceledException) { response = new(request.Id, Cancelled: true); }
    catch (Exception error) { response = new(request.Id, Error: error.Message); }
    finally { lock (cancellation) { requests.TryRemove(request.Id, out _); cancellation.Dispose(); } }
    try { await Send(response); }
    catch (InvalidOperationException error) { await Send(new(request.Id, Error: error.Message)); }
}
async Task<WorkerResponse> Drop(WorkerRequest request)
{
    await backend.DropAsync(request.Handle ?? throw new ArgumentException("Missing handle."));
    return new(request.Id);
}
async Task<WorkerResponse> Validate(WorkerRequest request, CancellationToken token)
{
    await backend.ValidateAsync(request.Query ?? throw new ArgumentException("Missing query."), token);
    return new(request.Id);
}
