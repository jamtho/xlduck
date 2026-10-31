// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Diagnostics;
using System.Text.Json;
using XlDuck.ExecutionPrototype;

var workerPath = Path.GetFullPath(args.Length > 0 ? args[0] : "prototypes/Execution/Worker/bin/Debug/net8.0/Worker.exe");
var measurements = new List<object>();
foreach (var mode in new[] { "inproc", "worker" })
{
    Session Create() => new(() => mode == "worker" ? new WorkerBackend(workerPath) : new LocalBackend());
    await using var session = Create();
    var watch = Stopwatch.StartNew();
    var source = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT $value AS value", Value: Atom.From(1d)));
    var sourceData = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: source.Handle, Limit: 1));
    if (sourceData.Rows![0][0] != Atom.From(1d)) throw new Exception("Numeric parameter lost its type.");
    var startupMs = watch.Elapsed.TotalMilliseconds;
    watch.Restart();
    var handle = source.Handle;
    double engineMs = 0;
    for (var i = 0; i < 20; i++)
    {
        var result = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT value+1 AS value FROM __input", handle));
        handle = result.Handle;
        engineMs += result.EngineMs;
    }
    var chainMs = watch.Elapsed.TotalMilliseconds;
    var leaf = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: handle, Limit: 1));
    if ((double)leaf.Rows![0][0].ToObject() != 21d) throw new Exception("Chain produced wrong value.");
    var beforeReuse = session.Executions;
    var repeated = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT $value AS value", Value: Atom.From(1d)));
    if (repeated.Handle != source.Handle || session.Executions != beforeReuse) throw new Exception("Reuse failed.");
    var sameSqlText = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT $value AS value", Value: Atom.From("1")));
    var sameSqlTextData = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: sameSqlText.Handle, Limit: 1));
    if (sameSqlText.Handle == source.Handle || sameSqlTextData.Rows![0][0] != Atom.From("1"))
        throw new Exception("Typed request identities collided.");

    var text = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT typeof($value) AS type", Value: Atom.From("1")));
    var textValue = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: text.Handle, Limit: 1));
    if ((string)textValue.Rows![0][0].ToObject() != "VARCHAR") throw new Exception("Typed text lost.");

    using var cancellation = new CancellationTokenSource();
    var slow = session.GetAsync(new(Guid.NewGuid(), "query", "SELECT sum(sin(i::DOUBLE)) FROM range(1000000000) t(i)"), cancellation.Token);
    await Task.Delay(100);
    watch.Restart();
    cancellation.Cancel();
    try { await slow; throw new Exception("Cancellation did not cancel."); }
    catch (OperationCanceledException) { }
    var cancellationMs = watch.Elapsed.TotalMilliseconds;
    var recovery = await session.GetAsync(new(Guid.NewGuid(), "query", "SELECT 42 AS value"));
    var recovered = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: recovery.Handle, Limit: 1));
    if ((double)recovered.Rows![0][0].ToObject() != 42d) throw new Exception("Failed recovery after cancellation.");

    await using var other = Create();
    try
    {
        await other.GetAsync(new(Guid.NewGuid(), "rows", Input: source.Handle));
        throw new Exception("A different session accepted a foreign handle.");
    }
    catch (Exception ex) when (ex.Message.Contains("does not belong")) { }
    var capture = Enumerable.Range(0, 10000).Select(i => new[] { Atom.From((double)i) }).ToArray();
    watch.Restart();
    var captured = await session.GetAsync(new(Guid.NewGuid(), "capture", Capture: capture));
    var captureMs = watch.Elapsed.TotalMilliseconds;
    if (captured.RowCount != 10000) throw new Exception("Capture row count mismatch.");
    watch.Restart();
    var output = await session.GetAsync(new(Guid.NewGuid(), "rows", Input: captured.Handle, Limit: 10000));
    var outputMs = watch.Elapsed.TotalMilliseconds;
    if (output.Rows!.Length != 10000) throw new Exception("Output row count mismatch.");
    measurements.Add(new { mode, startupMs, chainMs, engineMs, cancellationMs, captureMs, outputMs });
    Console.WriteLine($"PASS {mode}: chain, reuse, typed input, cancellation/recovery, isolation, capture/output.");
}
Console.WriteLine(JsonSerializer.Serialize(measurements, new JsonSerializerOptions { WriteIndented = true }));
Directory.CreateDirectory("artifacts/execution");
File.WriteAllText("artifacts/execution/core.json", JsonSerializer.Serialize(measurements, new JsonSerializerOptions { WriteIndented = true }));

await using (var crashingWorker = new WorkerBackend(workerPath))
{
    var pending = crashingWorker.ExecuteAsync(new(Guid.NewGuid(), "query",
        "SELECT sum(sin(i::DOUBLE)) FROM range(1000000000) t(i)"), CancellationToken.None);
    await Task.Delay(100);
    crashingWorker.SimulateCrash();
    try { await pending.WaitAsync(TimeSpan.FromSeconds(5)); throw new Exception("Worker termination was not reported."); }
    catch (IOException) { Console.WriteLine("PASS abrupt worker exit fails outstanding requests without terminating the host."); }
    try
    {
        await crashingWorker.ExecuteAsync(new(Guid.NewGuid(), "query", "SELECT 1"), CancellationToken.None)
            .WaitAsync(TimeSpan.FromSeconds(5));
        throw new Exception("Request after worker exit was not rejected.");
    }
    catch (IOException) { Console.WriteLine("PASS requests after worker exit fail promptly."); }
}

var controlled = new ControlledBackend();
await using (var sharing = new Session(() => controlled))
{
    using var firstCancellation = new CancellationTokenSource();
    var request = new Request(Guid.NewGuid(), "query", "shared");
    var first = sharing.GetAsync(request, firstCancellation.Token);
    var second = sharing.GetAsync(request with { Id = Guid.NewGuid() });
    await controlled.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
    firstCancellation.Cancel();
    try { await first; throw new Exception("First subscriber was not cancelled."); }
    catch (OperationCanceledException) { }
    if (controlled.Token.IsCancellationRequested) throw new Exception("Leaving subscriber cancelled shared work.");
    controlled.Release.TrySetResult(new(request.Id, "controlled-result"));
    if ((await second).Handle != "controlled-result" || sharing.Executions != 1)
        throw new Exception("Shared work was not preserved.");
    Console.WriteLine("PASS one cancelled subscriber does not cancel another subscriber's request.");
}

sealed class ControlledBackend : IBackend
{
    public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource<Result> Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public CancellationToken Token { get; private set; }
    public async Task<Result> ExecuteAsync(Request request, CancellationToken cancellationToken)
    {
        Token = cancellationToken;
        Started.TrySetResult();
        return await Release.Task.WaitAsync(cancellationToken);
    }
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
}
