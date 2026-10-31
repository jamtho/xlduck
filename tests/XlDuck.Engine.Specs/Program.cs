// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Diagnostics;
using System.Globalization;
using System.Numerics;
using System.Text.Json;
using XlDuck.Core;
using XlDuck.Engine;
using XlDuck.Client;

// Private child fixture: complete the handshake, then stop consuming requests.
if (args.Length == 4 && args[0] == "--session" && args[2] == "--spill")
{
    await WorkerProtocol.WriteAsync(Console.OpenStandardOutput(), new WorkerResponse(Guid.Empty, Ready: true));
    await Task.Delay(Timeout.Infinite);
    return 0;
}

int passed = 0, failed = 0;
var workerMode = args.Contains("--worker");
var workerPath = Path.GetFullPath("XlDuck.Worker/bin/Debug/net8.0/XlDuck.Worker.exe");
Console.WriteLine(workerMode ? "Worker transport and engine" : "In-process engine");
async Task Check(string name, Func<WorkbookSession, Task> test)
{
    DuckDbBackend? engine = null;
    WorkerBackend? worker = null;
    await using var session = new WorkbookSession(id => workerMode ? worker = new(id, workerPath) : engine = new(id), new(HistoricalResults: 0));
    try
    {
        await test(session).WaitAsync(TimeSpan.FromSeconds(15));
        await session.WaitForIdleAsync();
        var stored = worker is null ? engine?.TableCount ?? 0 : await worker.GetStoredTableCountAsync();
        if (session.Statistics.Nodes != 0 || stored != 0)
            throw new Exception($"Leaked {session.Statistics.Nodes} nodes / {stored} stored tables.");
        passed++; Console.WriteLine($"PASS {name}");
    }
    catch (Exception error) { failed++; Console.WriteLine($"FAIL {name}: {error.GetType().Name}: {error.Message}"); }
}
void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}."); }
async Task<TablePage> Query(WorkbookSession session, string sql, params QueryArgument[] arguments)
{
    using var lease = session.Acquire(ComputationRequest.Query(sql, arguments));
    await lease.Completion;
    return await session.ReadAsync(lease.Handle, new(ReadPurpose.Output));
}
async Task Reject(Func<Task> action, string contains)
{
    try { await action(); } catch (Exception ex) when (ex.Message.Contains(contains, StringComparison.OrdinalIgnoreCase)) { return; }
    throw new Exception($"Expected an error containing '{contains}'.");
}

await Check("typed scalar binding preserves numeric-looking text, apostrophes, empty text and following arguments", async session =>
{
    var page = await Query(session, "SELECT typeof(?) AS t, ? AS v, ? AS apostrophe, ? AS empty, ? AS n",
        QueryArgument.Scalar("00123"), QueryArgument.Scalar("00123"), QueryArgument.Scalar("'hello"), QueryArgument.Scalar(""), QueryArgument.Scalar(7d));
    Equal("VARCHAR", page.Rows[0][0].Text); Equal("00123", page.Rows[0][1].Text);
    Equal("'hello", page.Rows[0][2].Text); Equal("", page.Rows[0][3].Text); Equal(7d, page.Rows[0][4].ToDisplayValue());
});
await Check("question marks inside literals and comments are not bound", async session =>
{
    var page = await Query(session, "SELECT '?' AS literal, ? AS value -- why?", QueryArgument.Scalar(7d));
    Equal("?", page.Rows[0][0].Text); Equal(7d, page.Rows[0][1].ToDisplayValue());
});
await Check("captures keep quoted names, unique headers, booleans and nulls", async session =>
{
    var data = new CaptureData(["a", "a", "a_2", "a\"b", "flag"],
        [[CellValue.Number(1), CellValue.Number(2), CellValue.String("001"), CellValue.String("'x"), CellValue.From(true)],
         [CellValue.Empty, CellValue.Null, CellValue.String(""), CellValue.Null, CellValue.From(false)]]);
    using var captured = session.Acquire(ComputationRequest.CaptureRange(data));
    await captured.Completion;
    var page = await session.ReadAsync(captured.Handle, new(ReadPurpose.Output));
    Equal("a,a_3,a_2,a\"b,flag", string.Join(',', page.Columns.Select(c => c.Name)));
    Equal("001", page.Rows[0][2].Text); Equal("'x", page.Rows[0][3].Text);
    Equal(ValueKind.Null, page.Rows[1][0].Kind); Equal("", page.Rows[1][2].Text);
    Equal(true, page.Rows[0][4].ToDisplayValue());
});
await Check("mixed captures fail clearly and clean up", async session =>
{
    using var captured = session.Acquire(ComputationRequest.CaptureRange(new(["v"], [[CellValue.Number(1)], [CellValue.String("1")]])));
    await Reject(async () => await captured.Completion, "mixes value types");
});
await Check("capture ingestion supports 10000 rows with native appender", async session =>
{
    var rows = Enumerable.Range(0, 10000).Select(i => ImmutableArray.Create(CellValue.Number(i))).ToImmutableArray();
    using var capture = session.Acquire(ComputationRequest.CaptureRange(new(["v"], rows)));
    await capture.Completion;
    var page = await Query(session, "SELECT sum(v)::BIGINT AS total FROM ?", QueryArgument.Reference(capture.Handle));
    Equal(49995000d, page.Rows[0][0].ToDisplayValue());
});
await Check("large integers, decimal precision and nested nulls survive output", async session =>
{
    var page = await Query(session, "SELECT 9007199254740993::HUGEINT AS i, 12345678901234567890.123456789012345678::DECIMAL(38,18) AS d, [1,NULL,3] AS list, {'a':NULL,'b':9007199254740993::HUGEINT} AS nested");
    Equal("9007199254740993", page.Rows[0][0].ToDisplayValue());
    Equal("12345678901234567890.123456789012345678", page.Rows[0][1].ToDisplayValue());
    Equal("[1,null,3]", page.Rows[0][2].Text);
    using var nested = JsonDocument.Parse(page.Rows[0][3].Text);
    Equal(JsonValueKind.Null, nested.RootElement.GetProperty("a").ValueKind);
    Equal("9007199254740993", nested.RootElement.GetProperty("b").GetRawText());
});
await Check("dates, times and fractional intervals have explicit invariant representations", async session =>
{
    var previous = CultureInfo.CurrentCulture;
    try
    {
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
        var page = await Query(session, "SELECT DATE '2024-01-15' AS d, TIME '12:00:00' AS t, TIMESTAMP '2024-01-15 12:30:00' AS ts, INTERVAL '0.5 seconds' AS duration, 1.25::DOUBLE AS n");
        Equal(new DateTime(2024, 1, 15).ToOADate(), page.Rows[0][0].ToDisplayValue());
        Equal(0.5, page.Rows[0][1].ToDisplayValue());
        Equal(new DateTime(2024, 1, 15, 12, 30, 0).ToOADate(), page.Rows[0][2].ToDisplayValue());
        Equal("00:00:00.5", page.Rows[0][3].Text); Equal("1.25", page.Rows[0][4].Text);
    }
    finally { CultureInfo.CurrentCulture = previous; }
});
await Check("typed date, boolean and large-integer parameters bind correctly", async session =>
{
    var page = await Query(session, "SELECT ? AS d, ? AS flag, ? AS huge", QueryArgument.Scalar(new DateOnly(2024, 1, 15)), QueryArgument.Scalar(true), QueryArgument.Scalar(BigInteger.Parse("900719925474099312345")));
    Equal(ValueKind.Date, page.Rows[0][0].Kind); Equal(true, page.Rows[0][1].ToDisplayValue());
    Equal("900719925474099312345", page.Rows[0][2].Text);
});
await Check("38-digit decimal parameters never pass through rounded CLR decimal values", async session =>
{
    var exact = new CellValue(ValueKind.Decimal, "0.12345678901234567890123456789012345678");
    var page = await Query(session, "SELECT ? AS value", new QueryArgument(exact));
    Equal(ValueKind.Decimal, page.Rows[0][0].Kind);
    Equal(exact.Text, page.Rows[0][0].Text);
});
await Check("fragment composition and plots retain their actual table dependencies", async session =>
{
    var table = session.Acquire(ComputationRequest.Query("SELECT i AS n FROM range(5) t(i)"));
    await table.Completion;
    var fragment = session.Acquire(ComputationRequest.Fragment("SELECT * FROM ? WHERE n > ?", QueryArgument.Reference(table.Handle), QueryArgument.Scalar(2d)));
    await fragment.Completion;
    table.Dispose();
    using var plot = session.Acquire(ComputationRequest.Plot(fragment.Handle, "bar"));
    await plot.Completion;
    fragment.Dispose();
    using var data = session.AcquireData(plot.Handle);
    var first = await session.ReadAsync(data.Handle, new(ReadPurpose.Preview));
    var again = await session.ReadAsync(data.Handle, new(ReadPurpose.Preview));
    Equal(2L, first.TotalRows); Equal(3d, first.Rows[0][0].ToDisplayValue());
    Equal(first.Rows[0][0], again.Rows[0][0]);
});
await Check("output refuses truncation and scalar shape mismatches; previews label sampling", async session =>
{
    using var table = session.Acquire(ComputationRequest.Query("SELECT * FROM range(100)"));
    await table.Completion;
    await Reject(async () => await session.ReadAsync(table.Handle, new(ReadPurpose.Output, MaxRows: 10)), "exceeds the output limit");
    await Reject(async () => await session.ReadAsync(table.Handle, new(ReadPurpose.Scalar)), "exactly one row");
    var sample = await session.ReadAsync(table.Handle, new(ReadPurpose.Preview, MaxRows: 10));
    Equal(true, sample.IsSample); Equal(100L, sample.TotalRows); Equal(10, sample.Rows.Length);
});
await Check("statement parsing rejects mutations before execution", async session =>
{
    foreach (var sql in new[] { "CREATE TABLE forbidden(i INTEGER)", "SET threads=8", "COPY (SELECT 1) TO 'forbidden.csv'", "DELETE FROM missing", "INSTALL httpfs" })
    {
        using var request = session.Acquire(ComputationRequest.Query(sql));
        await Reject(async () => await request.Completion, "read-and-compute SELECT");
    }
    var page = await Query(session, "SELECT count(*) AS n FROM duckdb_tables() WHERE table_name = 'forbidden'");
    Equal(0d, page.Rows[0][0].ToDisplayValue());
});
await Check("native cancellation recovers without leaked partial tables", async session =>
{
    var slow = session.Acquire(ComputationRequest.Query("SELECT sum(sin(i::DOUBLE)) FROM range(1000000000) t(i)"));
    await Task.Delay(100);
    slow.Dispose();
    try { await slow.Completion; throw new Exception("Expected cancellation."); } catch (OperationCanceledException) { }
    var recovery = await Query(session, "SELECT 42 AS answer");
    Equal(42d, recovery.Rows[0][0].ToDisplayValue());
});
await Check("repeated edits retire physical tables as well as cache entries", async session =>
{
    for (var i = 0; i < 50; i++)
    {
        using (var table = session.Acquire(ComputationRequest.Query("SELECT ? AS n", QueryArgument.Scalar(i)))) await table.Completion;
        await session.WaitForIdleAsync();
    }
});
if (workerMode)
{
    try
    {
        WorkerBackend? worker = null;
        var session = new WorkbookSession(id => worker = new(id, workerPath), new(HistoricalResults: 0));
        await using (session)
        {
            using var retained = session.Acquire(ComputationRequest.Query("SELECT 1"));
            await retained.Completion;
            using var slow = session.Acquire(ComputationRequest.Query("SELECT sum(sin(i::DOUBLE)) FROM range(1000000000) t(i)"));
            // Give the private child a chance to start work; this test deliberately terminates that owned process.
            await Task.Delay(100);
            using var ownedProcess = Process.GetProcessById(worker!.ProcessId);
            File.WriteAllText(Path.Combine(worker.StoragePath, "cleanup-probe"), "owned test file");
            ownedProcess.Kill(entireProcessTree: true);
            await Reject(async () => await slow.Completion.WaitAsync(TimeSpan.FromSeconds(5)), "worker");
            using var afterExit = session.Acquire(ComputationRequest.Query("SELECT 2"));
            await Reject(async () => await afterExit.Completion.WaitAsync(TimeSpan.FromSeconds(5)), "worker");
        }
        Equal(false, Directory.Exists(worker!.StoragePath));
        passed++; Console.WriteLine("PASS worker exit fails pending/new requests and disposal removes owned spill storage");
    }
    catch (Exception error) { failed++; Console.WriteLine($"FAIL worker failure cleanup: {error}"); }
    foreach (var cancel in new[] { true, false })
    {
        WorkerBackend? stalled = null;
        try
        {
            var id = Guid.NewGuid();
            stalled = new WorkerBackend(id, Environment.ProcessPath!);
            using var cancellation = new CancellationTokenSource();
            var payload = new PreparedQuery("SELECT '" + new string('x', 2 * 1024 * 1024) + "'", [], []);
            var pending = stalled.ExecuteAsync(new(new(id, Guid.NewGuid(), ResultKind.Table), payload), cancellation.Token);
            await Task.Delay(200);
            if (cancel) cancellation.Cancel();
            else await stalled.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(8));
            try { await pending.WaitAsync(TimeSpan.FromSeconds(8)); throw new Exception("Expected blocked transport to fail."); }
            catch (Exception error) when (error is IOException or OperationCanceledException or ObjectDisposedException) { }
            await stalled.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(8));
            Equal(false, Directory.Exists(stalled.StoragePath));
            passed++; Console.WriteLine($"PASS {(cancel ? "cancellation" : "disposal")} remains bounded when an owned worker stops reading its pipe");
        }
        catch (Exception error) { failed++; Console.WriteLine($"FAIL blocked worker transport: {error}"); }
        finally
        {
            if (stalled is not null)
            {
                // A failing test may leave its deliberately wedged private child; never touch other processes.
                try { using var child = Process.GetProcessById(stalled.ProcessId); if (!child.HasExited) child.Kill(); }
                catch (ArgumentException) { }
                await stalled.DisposeAsync();
            }
        }
    }
}
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
