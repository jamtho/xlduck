// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Concurrent;
using System.Text.Json;
using XlDuck.ExecutionPrototype;

await using var backend = new LocalBackend();
using var output = new SemaphoreSlim(1);
var requests = new ConcurrentDictionary<Guid, CancellationTokenSource>();
var running = new ConcurrentBag<Task>();
while (await Console.In.ReadLineAsync() is { } line)
{
    var request = JsonSerializer.Deserialize<Request>(line) ?? throw new IOException("Invalid request.");
    if (request.Operation == "cancel")
    {
        if (requests.TryGetValue(request.Id, out var pending))
            lock (pending) { if (requests.ContainsKey(request.Id)) pending.Cancel(); }
        continue;
    }
    var cancellation = new CancellationTokenSource();
    requests[request.Id] = cancellation;
    running.Add(RespondAsync(request, cancellation));
}
foreach (var pair in requests)
    lock (pair.Value) { if (requests.ContainsKey(pair.Key)) pair.Value.Cancel(); }
await Task.WhenAll(running);

async Task RespondAsync(Request request, CancellationTokenSource cancellation)
{
    Result result;
    try { result = await backend.ExecuteAsync(request, cancellation.Token); }
    catch (Exception ex) { result = new(request.Id, Error: ex.Message); }
    finally
    {
        lock (cancellation) { requests.TryRemove(request.Id, out _); cancellation.Dispose(); }
    }
    await output.WaitAsync();
    try { await Console.Out.WriteLineAsync(JsonSerializer.Serialize(result)); }
    finally { output.Release(); }
}
