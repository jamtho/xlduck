// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Globalization;
using System.Text.Json;
using XlDuck.Core;
using System.Diagnostics;

if (args is ["--orphan-spill"])
{
    var orphan = new SessionDirectory();
    Console.WriteLine(orphan.Path);
    return 0; // Deliberately omit Dispose, simulating process exit before cleanup.
}

int passed = 0, failed = 0;
async Task Check(string name, Func<Task> test)
{
    try { await test().WaitAsync(TimeSpan.FromSeconds(10)); passed++; Console.WriteLine($"PASS {name}"); }
    catch (Exception error) { failed++; Console.WriteLine($"FAIL {name}: {error}"); }
}
void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}."); }
void Throws<T>(Action action) where T : Exception { try { action(); } catch (T) { return; } throw new Exception($"Expected {typeof(T).Name}."); }
PreparedQuery Prepare(string sql, params QueryArgument[] args) => SqlTemplate.Prepare(ComputationRequest.Query(sql, args), _ => null);
CaptureData Capture(CellValue value) => new(["value"], [[value]]);

await Check("compact handles resolve only live results in their own workbook session", async () =>
{
    await using var session = new WorkbookSession(_ => new FakeBackend(), new(HistoricalResults: 0));
    await using var other = new WorkbookSession(_ => new FakeBackend());
    var lease = session.Acquire(ComputationRequest.Query("SELECT 1")); await lease.Completion;
    var text = lease.Handle.CellText;
    Equal(37, text.Length); Equal(false, text.Contains(session.Id.ToString("N")));
    Equal(lease.Handle, session.ResolveCellHandle(text));
    Throws<ArgumentException>(() => other.ResolveCellHandle(text));
    Throws<ArgumentException>(() => session.ResolveCellHandle(text.Replace("/table/", "/plot/")));
    Throws<ArgumentException>(() => session.ResolveCellHandle(text + "="));
    Throws<ArgumentException>(() => session.ResolveCellHandle(lease.Handle.ToString()));
    using (var fragment = session.Acquire(ComputationRequest.Fragment("SELECT * FROM ?", QueryArgument.Reference(lease.Handle))))
    {
        await fragment.Completion; Equal(fragment.Handle, session.ResolveCellHandle(fragment.Handle.CellText));
        using var plot = session.Acquire(ComputationRequest.Plot(lease.Handle, "bar")); await plot.Completion;
        Equal(plot.Handle, session.ResolveCellHandle(plot.Handle.CellText));
    }
    lease.Dispose(); await session.WaitForIdleAsync();
    Throws<ArgumentException>(() => session.ResolveCellHandle(text));
});
await Check("spill recovery preserves active hosts/workers and unmarked directories", async () =>
{
    using var active = new SessionDirectory();
    var start = new ProcessStartInfo(Environment.ProcessPath!) { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
    if (System.IO.Path.GetFileNameWithoutExtension(Environment.ProcessPath) == "dotnet") start.ArgumentList.Add(typeof(Program).Assembly.Location);
    start.ArgumentList.Add("--orphan-spill");
    using var child = Process.Start(start)!;
    var orphan = (await child.StandardOutput.ReadLineAsync())!;
    await child.WaitForExitAsync();
    Equal(0, child.ExitCode);
    Equal(System.IO.Path.GetDirectoryName(active.Path), System.IO.Path.GetDirectoryName(orphan));
    using (var worker = SessionDirectory.HoldWorkerOwnership(orphan))
    {
        SessionDirectory.ReclaimOrphans(); Equal(true, Directory.Exists(orphan)); Equal(true, Directory.Exists(active.Path));
    }
    SessionDirectory.ReclaimOrphans(); Equal(false, Directory.Exists(orphan)); Equal(true, Directory.Exists(active.Path));
    var foreign = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(active.Path)!, "xlduck-" + Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(foreign);
    try { SessionDirectory.ReclaimOrphans(); Equal(true, Directory.Exists(foreign)); }
    finally { Directory.Delete(foreign); }
});
await Check("typed values preserve text, apostrophes, blank distinctions and exact integers", () =>
{
    Equal("00123", CellValue.From("00123").Text);
    Equal("'hello", CellValue.From("'hello").Text);
    Equal(false, CellValue.Empty == CellValue.Null);
    Equal(false, CellValue.String("") == CellValue.Null);
    Equal("9007199254740993", CellValue.From(9007199254740993L).ToDisplayValue());
    Equal("1.25", CellValue.From(1.25m).ToDisplayValue());
    Equal(true, CellValue.From(true).ToDisplayValue());
    Throws<ArgumentException>(() => CellValue.Number(double.NaN));
    Throws<ArgumentException>(() => new CellValue(ValueKind.Error, "#N/A").Validate());
    return Task.CompletedTask;
});
await Check("capture and query identities preserve typed values and dimensions", () =>
{
    var values = new[] { CellValue.Number(1), CellValue.String("1"), CellValue.Null, CellValue.Empty, CellValue.String("") };
    Equal(5, values.Select(v => ComputationRequest.CaptureRange(Capture(v)).Identity()).Distinct().Count());
    Equal(5, values.Select(v => ComputationRequest.Query("SELECT ?", new QueryArgument(v)).Identity()).Distinct().Count());
    var data = Capture(CellValue.Number(1));
    var copy = JsonSerializer.Deserialize<CaptureData>(JsonSerializer.Serialize(data))!;
    Equal(ComputationRequest.CaptureRange(data).Identity(), ComputationRequest.CaptureRange(copy).Identity());
    return Task.CompletedTask;
});
await Check("placeholder parsing respects quotes, escaped strings and nested comments", () =>
{
    var prepared = Prepare("SELECT '?' AS \"?\", E'it\\'s ?' AS e, $$?$$ AS d, $tag$?$tag$ AS t, ? /* ? /* ? */ */ -- ?\n, ?; -- trailing ?",
        QueryArgument.Scalar(""), QueryArgument.Scalar(7d));
    Equal(2, prepared.Parameters.Length);
    Equal(CellValue.String(""), prepared.Parameters[0].Value);
    Equal(CellValue.Number(7), prepared.Parameters[1].Value);
    Equal(true, prepared.Sql.Contains("$__xl0", StringComparison.Ordinal));
    Throws<ArgumentException>(() => Prepare("SELECT ?", QueryArgument.Scalar(1), QueryArgument.Scalar(2)));
    Throws<ArgumentException>(() => Prepare("SELECT ?"));
    Throws<ArgumentException>(() => Prepare("SELECT 1; DELETE FROM x"));
    Throws<ArgumentException>(() => Prepare("SELECT 'unterminated"));
    Throws<ArgumentException>(() => Prepare("SELECT 1 /* unterminated"));
    Throws<ArgumentException>(() => Prepare("SELECT $name"));
    return Task.CompletedTask;
});
await Check("headers are unique without stealing explicit names; identifiers are escaped", () =>
{
    var names = CaptureData.UniqueHeaders(["a", "a", "a_2", "A", ""]);
    Equal("a,a_3,a_2,A_4,Column5", string.Join(',', names));
    Equal("\"a\"\"b\"", SqlTemplate.QuoteIdentifier("a\"b"));
    return Task.CompletedTask;
});
await Check("values and identities do not change with locale", () =>
{
    var request = ComputationRequest.Query("SELECT ?", QueryArgument.Scalar(1.25));
    var before = request.Identity();
    var previous = CultureInfo.CurrentCulture;
    try
    {
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
        Equal(before, ComputationRequest.Query("SELECT ?", QueryArgument.Scalar(1.25)).Identity());
        Equal("1.25", CellValue.Number(1.25).Text);
        Equal(1.25, CellValue.Number(1.25).ToDisplayValue());
    }
    finally { CultureInfo.CurrentCulture = previous; }
    return Task.CompletedTask;
});
await Check("two leases share a request and one may leave without cancelling the other", async () =>
{
    var backend = new FakeBackend(block: true);
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 0));
    var request = ComputationRequest.Query("SELECT 1");
    var first = session.Acquire(request);
    using var second = session.Acquire(request);
    Equal(first.Handle, second.Handle);
    await backend.Started.Task;
    first.Dispose();
    Equal(false, backend.Token.IsCancellationRequested);
    backend.Continue.TrySetResult();
    await second.Completion;
    Equal(1, backend.Executions);
    using var third = session.Acquire(request);
    Equal(second.Handle, third.Handle);
});
await Check("last owner leaving cancels pending work and releases its dependencies", async () =>
{
    var backend = new FakeBackend(block: true);
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 0));
    var lease = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await backend.Started.Task;
    lease.Dispose();
    try { await lease.Completion; throw new Exception("Expected cancellation."); } catch (OperationCanceledException) { }
    await session.WaitForIdleAsync();
    Equal(0, session.Statistics.Nodes);
});
await Check("fragments retain tables; preparation failures acquire no leaked references", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 0));
    var table = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await table.Completion;
    var before = session.Statistics;
    Throws<ArgumentException>(() => session.Acquire(ComputationRequest.Query("SELECT * FROM ? WHERE x=?", QueryArgument.Reference(table.Handle))));
    Equal(before.ReferencedNodes, session.Statistics.ReferencedNodes);
    var fragment = session.Acquire(ComputationRequest.Fragment("SELECT * FROM ? WHERE x=? -- ?", QueryArgument.Reference(table.Handle), QueryArgument.Scalar(1)));
    await fragment.Completion;
    table.Dispose();
    using (var query = session.Acquire(ComputationRequest.Query("SELECT * FROM ?", QueryArgument.Reference(fragment.Handle))))
    {
        await query.Completion;
        Equal(true, backend.Last!.Query!.Sql.Contains(SqlTemplate.QuoteIdentifier(table.Handle.TableName), StringComparison.Ordinal));
        Equal(1, backend.Last.Query.Parameters.Length);
    }
    fragment.Dispose();
    await session.WaitForIdleAsync();
    Equal(0, session.Statistics.Nodes);
    Equal(0, backend.Tables.Count);
});
await Check("plots retain one materialization of a fragment", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 0));
    var fragment = session.Acquire(ComputationRequest.Fragment("SELECT random()"));
    await fragment.Completion;
    var plot = session.Acquire(ComputationRequest.Plot(fragment.Handle, "line"));
    await plot.Completion;
    fragment.Dispose();
    for (var i = 0; i < 3; i++)
    {
        using var data = session.AcquireData(plot.Handle);
        await session.ReadAsync(data.Handle, new(ReadPurpose.Preview));
    }
    Equal(1, backend.Executions);
    plot.Dispose();
    await session.WaitForIdleAsync();
    Equal(0, backend.Tables.Count);
});
await Check("unused history is bounded and live results survive cache pressure", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 3));
    using var live = session.Acquire(ComputationRequest.Query("SELECT 999"));
    await live.Completion;
    for (var i = 0; i < 100; i++)
    {
        using (var lease = session.Acquire(ComputationRequest.Query("SELECT ?", QueryArgument.Scalar(i)))) await lease.Completion;
        await session.WaitForIdleAsync();
    }
    Equal(4, session.Statistics.Nodes);
    Equal(4, backend.Tables.Count);
    using var again = session.Acquire(ComputationRequest.Query("SELECT 999"));
    Equal(live.Handle, again.Handle);
});
await Check("admission is bounded and foreign session handles are rejected", async () =>
{
    var backend = new FakeBackend(block: true);
    await using var session = new WorkbookSession(_ => backend, new(MaxPending: 1));
    using var first = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await backend.Started.Task;
    Throws<InvalidOperationException>(() => session.Acquire(ComputationRequest.Query("SELECT 2")));
    backend.Continue.TrySetResult();
    await first.Completion;
    await using var other = new WorkbookSession(_ => new FakeBackend());
    Throws<ArgumentException>(() => other.Acquire(ComputationRequest.Query("SELECT * FROM ?", QueryArgument.Reference(first.Handle))));
});
await Check("closing a session cancels work, releases storage and rejects old handles", async () =>
{
    var backend = new FakeBackend(block: true);
    var session = new WorkbookSession(_ => backend);
    using var lease = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await backend.Started.Task;
    await session.DisposeAsync();
    Equal(true, backend.Disposed);
    Equal(true, lease.Completion.IsCanceled);
    Equal(0, session.Statistics.Nodes);
    Throws<ObjectDisposedException>(() => session.AcquireExisting(lease.Handle));
    await session.DisposeAsync();
});
await Check("unused fragments cannot keep an unbounded dependency graph in history", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 3));
    var previous = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await previous.Completion;
    for (var i = 0; i < 40; i++)
    {
        var next = session.Acquire(ComputationRequest.Fragment("SELECT * FROM ?", QueryArgument.Reference(previous.Handle)));
        await next.Completion;
        previous.Dispose(); previous = next;
    }
    previous.Dispose();
    await session.WaitForIdleAsync();
    Equal(1, session.Statistics.Nodes);
    Equal(1, session.Statistics.Historical);
});
await Check("failed requests stay stable until an explicit retry", async () =>
{
    var backend = new FakeBackend { FailNext = true };
    await using var session = new WorkbookSession(_ => backend, new(HistoricalResults: 0));
    var request = ComputationRequest.Query("SELECT 1");
    using var first = session.Acquire(request);
    try { await first.Completion; throw new Exception("Expected engine failure."); } catch (InvalidOperationException) { }
    using var reused = session.Acquire(request);
    Equal(first.Handle, reused.Handle); Equal(1, backend.Executions);
    using var retry = session.Acquire(request, retryFailed: true);
    await retry.Completion;
    Equal(false, retry.Handle == first.Handle); Equal(2, backend.Executions);
});
await Check("explicit retry invalidates failed fragment materializations for a later calculation", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend);
    using var fragment = session.Acquire(ComputationRequest.Fragment("SELECT 1")); await fragment.Completion;
    backend.FailNext = true;
    using var failed = session.AcquireData(fragment.Handle);
    try { await failed.Completion; throw new Exception("Expected failure"); } catch (InvalidOperationException) { }
    session.InvalidateFailuresForRetry();
    using var later = session.AcquireData(fragment.Handle); await later.Completion;
    Equal(false, failed.Handle == later.Handle);
    Equal(true, failed.Completion.IsFaulted);
    using var stable = session.AcquireData(fragment.Handle); Equal(later.Handle, stable.Handle);
});
await Check("output admission and input sizes are bounded", async () =>
{
    var backend = new FakeBackend { BlockReads = true };
    await using var session = new WorkbookSession(_ => backend, new(MaxPending: 1, MaxInputCharacters: 32, MaxCaptureCells: 2));
    Throws<ArgumentException>(() => session.Acquire(ComputationRequest.Query("SELECT ?", QueryArgument.Scalar(new string('x', 40)))));
    Throws<ArgumentException>(() => session.Acquire(ComputationRequest.CaptureRange(new(["x"], [[CellValue.Null], [CellValue.Null], [CellValue.Null]]))));
    using var table = session.Acquire(ComputationRequest.Query("SELECT 1"));
    await table.Completion;
    var first = session.ReadAsync(table.Handle, new(ReadPurpose.Output));
    await backend.ReadStarted.Task;
    Throws<InvalidOperationException>(() => session.ReadAsync(table.Handle, new(ReadPurpose.Preview)));
    backend.Continue.TrySetResult();
    await first;
});
await Check("worker framing preserves typed data and rejects oversized/truncated frames", async () =>
{
    using var stream = new MemoryStream();
    var request = new WorkerRequest(Guid.NewGuid(), WorkerOperation.Read, Handle: new(Guid.NewGuid(), Guid.NewGuid(), ResultKind.Table), Read: new(ReadPurpose.Preview));
    await WorkerProtocol.WriteAsync(stream, request);
    stream.Position = 0;
    Equal(request, await WorkerProtocol.ReadAsync<WorkerRequest>(stream));
    using var startupStream = new MemoryStream();
    await WorkerProtocol.WriteAsync(startupStream, new WorkerStartup(new()));
    startupStream.Position = 0;
    var startup = await WorkerProtocol.ReadAsync<WorkerStartup>(startupStream);
    Equal(0, startup!.Settings.S3.Length);
    Equal(false, startup.Settings.EnableExternalReads);
    using var oversized = new MemoryStream(BitConverter.GetBytes(WorkerProtocol.MaxFrameBytes + 1));
    try { await WorkerProtocol.ReadAsync<WorkerRequest>(oversized); throw new Exception("Expected size rejection."); } catch (IOException) { }
    using var truncated = new MemoryStream([1, 0]);
    try { await WorkerProtocol.ReadAsync<WorkerRequest>(truncated); throw new Exception("Expected truncation rejection."); } catch (EndOfStreamException) { }
});
await Check("fragment expansion is bounded before it can allocate an excessive SQL tree", async () =>
{
    await using var session = new WorkbookSession(_ => new FakeBackend(), new(HistoricalResults: 0, MaxInputCharacters: 1024));
    var leases = new List<ResultLease>();
    try
    {
        var current = session.Acquire(ComputationRequest.Fragment("SELECT 1"));
        leases.Add(current); await current.Completion;
        var rejected = false;
        for (var i = 0; i < 10; i++)
        {
            try { current = session.Acquire(ComputationRequest.Fragment("SELECT * FROM ? UNION ALL SELECT * FROM ?", QueryArgument.Reference(current.Handle), QueryArgument.Reference(current.Handle))); }
            catch (ArgumentException error) when (error.Message.Contains("Expanded SQL", StringComparison.Ordinal)) { rejected = true; break; }
            leases.Add(current); await current.Completion;
        }
        Equal(true, rejected);
    }
    finally { foreach (var lease in leases) lease.Dispose(); }
    await session.WaitForIdleAsync();
    Equal(0, session.Statistics.Nodes);
});
await Check("pause retains pending requests; resume reuses work and cancel preserves completed results", async () =>
{
    var backend = new FakeBackend();
    await using var session = new WorkbookSession(_ => backend);
    using var completed = session.Acquire(ComputationRequest.Query("SELECT 1")); await completed.Completion;
    session.SetPaused(true);
    using var pending = session.Acquire(ComputationRequest.Query("SELECT 2"));
    using var shared = session.Acquire(ComputationRequest.Query("SELECT 2"));
    Equal(pending.Handle, shared.Handle); Equal(false, pending.Completion.IsCompleted);
    Equal(1, backend.Executions);
    session.CancelPending();
    try { await pending.Completion; throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
    Equal(true, completed.Completion.IsCompletedSuccessfully);
    using var retry = session.Acquire(ComputationRequest.Query("SELECT 2"), retryFailed: true);
    Equal(false, retry.Handle == pending.Handle);
    session.SetPaused(false); await retry.Completion;
    Equal(2, backend.Executions);
    using var reused = session.Acquire(ComputationRequest.Query("SELECT 2")); Equal(retry.Handle, reused.Handle);
});
await Check("cancel interrupts output reads without destroying their source result", async () =>
{
    var backend = new FakeBackend { BlockReads = true };
    await using var session = new WorkbookSession(_ => backend);
    using var table = session.Acquire(ComputationRequest.Query("SELECT 1")); await table.Completion;
    var read = session.ReadAsync(table.Handle, new(ReadPurpose.Preview)); await backend.ReadStarted.Task;
    session.CancelPending();
    try { await read; throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
    backend.BlockReads = false;
    Equal(1L, (await session.ReadAsync(table.Handle, new(ReadPurpose.Output))).TotalRows);
});
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;

sealed class FakeBackend(bool block = false) : IExecutionBackend
{
    public Task ValidateAsync(PreparedQuery query, CancellationToken cancellationToken) => Task.CompletedTask;
    public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource Continue { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public CancellationToken Token;
    public int Executions;
    public bool Disposed;
    public bool FailNext;
    public bool BlockReads;
    public TaskCompletionSource ReadStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public BackendRequest? Last;
    public readonly System.Collections.Concurrent.ConcurrentDictionary<ResultHandle, bool> Tables = new();
    public async Task<ComputationResult> ExecuteAsync(BackendRequest request, CancellationToken cancellationToken)
    {
        Token = cancellationToken; Last = request;
        Interlocked.Increment(ref Executions);
        Started.TrySetResult();
        if (block) await Continue.Task.WaitAsync(cancellationToken);
        if (FailNext) { FailNext = false; throw new InvalidOperationException("Test engine failure."); }
        Tables[request.Output] = true;
        return new(request.Output, [new("x", "INTEGER")], 1);
    }
    public async Task<TablePage> ReadAsync(ResultHandle handle, ReadRequest request, CancellationToken cancellationToken)
    {
        ReadStarted.TrySetResult();
        if (BlockReads) await Continue.Task.WaitAsync(cancellationToken);
        return new([new("x", "INTEGER")], [[CellValue.From(1)]], 1, false);
    }
    public Task DropAsync(ResultHandle handle) { Tables.TryRemove(handle, out _); return Task.CompletedTask; }
    public ValueTask DisposeAsync() { Disposed = true; Tables.Clear(); return ValueTask.CompletedTask; }
}
