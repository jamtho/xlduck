// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using DuckDB.NET.Data;
using XlDuck.Client;
using XlDuck.Core;
using XlDuck.Engine;

var workerMode = args.Contains("--worker");
var network = args.Contains("--readers");
var extensionPath = Path.GetFullPath("artifacts/reader-extensions/v1.5.6/windows_amd64");
var workerPath = Path.GetFullPath("XlDuck.Worker/bin/Debug/net8.0/XlDuck.Worker.exe");
using var files = new SessionDirectory();
var csv = Path.Combine(files.Path, "sample.csv");
var json = Path.Combine(files.Path, "sample.json");
var parquet = Path.Combine(files.Path, "sample.parquet");
var forbidden = Path.Combine(files.Path, "forbidden.csv");
File.WriteAllText(csv, "id,name\n1,alice\n2,bob\n");
File.WriteAllText(json, "[{\"id\":1,\"name\":\"alice\"},{\"id\":2,\"name\":\"bob\"}]");
// Privileged fixture creation is separate from the worksheet path under test.
using (var fixture = new DuckDBConnection("DataSource=:memory:"))
{
    fixture.Open(); using var create = fixture.CreateCommand();
    create.CommandText = $"COPY (SELECT * FROM (VALUES (1, 'alice'), (2, 'bob')) t(id,name)) TO {Literal(parquet)} (FORMAT PARQUET)";
    create.ExecuteNonQuery();
}
int passed = 0, failed = 0;
Console.WriteLine($"Read-and-compute policy ({(workerMode ? "worker" : "direct engine")}, readers={network})");
WorkbookSession Session(ConnectionSettings? settings = null) => new(id => workerMode ? new WorkerBackend(id, workerPath, settings) : new DuckDbBackend(id, settings: settings), new(HistoricalResults: 0));
async Task Check(string name, Func<Task> test)
{
    try { await test().WaitAsync(TimeSpan.FromSeconds(25)); passed++; Console.WriteLine("PASS " + name); }
    catch (Exception error) { failed++; Console.WriteLine($"FAIL {name}: {error}"); }
}
void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Expected {expected}, got {actual}."); }
string Literal(string text) => "'" + text.Replace("'", "''", StringComparison.Ordinal) + "'";
async Task<TablePage> Query(WorkbookSession session, string sql, params QueryArgument[] arguments)
{
    using var lease = session.Acquire(ComputationRequest.Query(sql, arguments));
    await lease.Completion;
    return await session.ReadAsync(lease.Handle, new(ReadPurpose.Output));
}
async Task<string> Reject(Func<Task> action, string expected)
{
    try { await action(); }
    catch (Exception error) when (error.Message.Contains(expected, StringComparison.OrdinalIgnoreCase)) { return error.Message; }
    throw new Exception("Expected rejection containing: " + expected);
}
await Check("unenabled sessions compute but cannot read external data", async () =>
{
    await using var session = Session();
    Equal(3d, (await Query(session, "SELECT 1+2")).Rows[0][0].ToDisplayValue());
    await Reject(async () => await Query(session, "SELECT * FROM read_csv(?)", QueryArgument.Scalar(csv)), "disabled");
});
await Check("enabled local CSV, JSON and Parquet reads use typed parameters", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    foreach (var (reader, path) in new[] { ("read_csv", csv), ("read_json", json), ("read_parquet", parquet) })
    {
        var page = await Query(session, $"SELECT id, name FROM {reader}(?) ORDER BY id", QueryArgument.Scalar(path));
        Equal(2L, page.TotalRows); Equal("alice", page.Rows[0][1].Text);
    }
});
await Check("CTEs, windows, macros, lambdas and random computations remain available", async () =>
{
    await using var session = Session();
    var page = await Query(session, "WITH t AS (SELECT * FROM range(3) r(n)) SELECT sum(n) OVER (), list_sum(list_transform([1,2], x -> x+1)), date_add(DATE '2024-01-01', INTERVAL 1 DAY), random() FROM t");
    Equal(3L, page.TotalRows); Equal(5d, page.Rows[0][1].ToDisplayValue());
    Equal(6d, (await Query(session, "WITH RECURSIVE t(n) AS (SELECT 1 UNION ALL SELECT n+1 FROM t WHERE n<3) SELECT sum(n) FROM t")).Rows[0][0].ToDisplayValue());
    Equal(1d, (await Query(session, "WITH a AS (SELECT 1 n), b AS (SELECT * FROM a) SELECT * FROM b")).Rows[0][0].ToDisplayValue());
    Equal(1d, (await Query(session, "WITH z AS (SELECT 1 n), a AS (SELECT * FROM z) SELECT * FROM a")).Rows[0][0].ToDisplayValue());
    Equal(2d, (await Query(session, "WITH a AS (SELECT 1 n) SELECT * FROM (WITH a AS (SELECT 2 n) SELECT * FROM a) q")).Rows[0][0].ToDisplayValue());
    Equal(1d, (await Query(session, "PIVOT (SELECT 'a' k, 1 n) ON k IN ('a') USING sum(n)")).Rows[0][0].ToDisplayValue());
});
await Check("UNNEST works in SELECT expressions, recursively and in FROM without bypassing policy", async () =>
{
    await using var session = Session();
    Equal(3L, (await Query(session, "SELECT unnest([1,2,3]) AS n")).TotalRows);
    Equal(3L, (await Query(session, "SELECT * FROM unnest([1,2,3]) AS t(n)")).TotalRows);
    var nested = await Query(session, "SELECT unnest([{'a': 1, 'b': [2,3]}], recursive := true)");
    Equal("a", nested.Columns[0].Name); Equal("b", nested.Columns[1].Name);
    Equal(3L, (await Query(session, "SELECT unnest([[1,2],[3]], recursive := true) AS n")).TotalRows);
    await Reject(async () => await Query(session, "SELECT unnest([nextval('forbidden')])"), "read-and-compute SQL");
});
await Check("schema-only queries bind columns without evaluating data expressions", async () =>
{
    await using var session = Session();
    using var schema = session.Acquire(ComputationRequest.Query("SELECT error('must not execute') AS n, 1::BIGINT AS id; -- end") with { RowLimit = 0 });
    var result = await schema.Completion;
    Equal(0L, result.RowCount); Equal("BIGINT", result.Columns[1].Type);
    await Reject(async () => await Query(session, "DESCRIBE SELECT 1 n"), "DuckSchema");
    using var describe = session.Acquire(ComputationRequest.Fragment("DESCRIBE SELECT 1 n"));
    await Reject(async () => await describe.Completion, "DuckSchema");
    using var bad = session.Acquire(ComputationRequest.Query("SELECT missing FROM (SELECT 1 n)") with { RowLimit = 0 });
    try { await bad.Completion; throw new Exception("Expected binding error"); }
    catch (ComputationException error) { Equal("sql_binding", error.Failure.Code); }
});
await Check("exploration shares the worker and bounds results while preserving exact values", async () =>
{
    await using var session = Session(); using var explorer = new ExplorationSession(session);
    async Task<ExplorationStatus> Explore(ExplorationRequest request)
    {
        var status = explorer.Invoke(request);
        while (status.State is not ("succeeded" or "failed" or "cancelled"))
        {
            await Task.Delay(10); status = explorer.Invoke(new("poll", request.Id));
        }
        return status;
    }
    var sample = await Explore(new("query", "sample", "SELECT i, 9007199254740993::HUGEINT AS exact FROM range(100) t(i)", MaxRows: 3));
    Equal("succeeded", sample.State); Equal(3, sample.Result!.Rows.Length); Equal(true, sample.Result.HasMore);
    Equal("9007199254740993", sample.Result.Rows[0][1].Text);
    var overflow = await Explore(new("query", "large", "SELECT repeat('x', 31000) AS body"));
    Equal("output_limit", overflow.Failure!.Code); Equal(true, ExplorationSession.Serialize(overflow).Length < 30000);
    var validation = await Explore(new("validate", "validate", "SELECT not_bound FROM range(3)"));
    Equal("succeeded", validation.State); Equal("syntax_and_capabilities_only", validation.Result!.Validation);
    var binding = await Explore(new("schema", "bind", "SELECT not_bound FROM range(3)"));
    Equal("sql_binding", binding.Failure!.Code);
    var denied = await Explore(new("query", "denied", "SELECT unnest([nextval('forbidden')])"));
    Equal("sql_validation", denied.Failure!.Code);
    using var fragment = session.Acquire(ComputationRequest.Fragment("SELECT error('must not execute')")); await fragment.Completion;
    var inspect = await Explore(new("inspect", "recipe", Handle: fragment.CellText));
    Equal("succeeded", inspect.State); Equal(RequestKind.Fragment, inspect.Result!.Definition!.Kind);
    Equal(0, inspect.Result.Columns.Length);
});
await Check("resource budgets are supplied by the host and failures retain structured categories", async () =>
{
    await using var session = Session(new(Resources: new(128, 2, 1)));
    Equal("128.0 MiB", (await Query(session, "SELECT current_setting('memory_limit')")).Rows[0][0].Text);
    Equal("2.0 GiB", (await Query(session, "SELECT current_setting('max_temp_directory_size')")).Rows[0][0].Text);
    Equal(1d, (await Query(session, "SELECT current_setting('threads')")).Rows[0][0].ToDisplayValue());
    try { await Query(session, "SELECT list(n) FROM range(100000000) t(n)"); throw new Exception("Expected memory failure"); }
    catch (ComputationException error)
    {
        Equal("memory_limit", error.Failure.Code);
        Equal(true, error.Failure.Advice!.Contains("metadata"));
        Equal(false, error.Message.Contains("SET "));
    }
    try { await Query(session, "SELECT CAST('private' AS INTEGER)"); throw new Exception("Expected conversion failure"); }
    catch (ComputationException error) { Equal("data_conversion", error.Failure.Code); Equal(false, error.Message.Contains("private")); }
});
await Check("changed external data is reused until a new session is explicitly created", async () =>
{
    var changing = Path.Combine(files.Path, "changing.csv");
    File.WriteAllText(changing, "n\n1\n");
    var request = ComputationRequest.Query("SELECT * FROM read_csv(?)", QueryArgument.Scalar(changing));
    await using var first = Session(new(EnableExternalReads: true));
    using var retained = first.Acquire(request); await retained.Completion;
    File.WriteAllText(changing, "n\n2\n3\n");
    using var same = first.Acquire(request); Equal(retained.Handle, same.Handle);
    Equal(1L, (await first.ReadAsync(same.Handle, new(ReadPurpose.Output))).TotalRows);
    await using var restarted = Session(new(EnableExternalReads: true));
    using var fresh = restarted.Acquire(request); await fresh.Completion;
    Equal(2L, (await restarted.ReadAsync(fresh.Handle, new(ReadPurpose.Output))).TotalRows);
});
await Check("DDL, DML, export, settings and installation cannot create effects", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    foreach (var sql in new[] { "CREATE TABLE forbidden(i INT)", "CREATE SECRET x(TYPE s3)", "CREATE MACRO f() AS 1", "CREATE SEQUENCE s", "DELETE FROM missing", "UPDATE missing SET i=1", "INSERT INTO missing VALUES(1)", "SET threads=8", "PRAGMA enable_profiling", "INSTALL httpfs", "LOAD httpfs", "ATTACH 'forbidden.duckdb'", $"COPY (SELECT 1) TO {Literal(forbidden)}", "SELECT 1; CREATE TABLE forbidden(i INT)" })
        await Reject(async () => await Query(session, sql), sql.Contains(';') ? "one SQL statement" : "read-and-compute SELECT");
    Equal(false, File.Exists(forbidden));
    Equal(0d, (await Query(session, "SELECT count(*) FROM duckdb_tables() WHERE table_name='forbidden'")).Rows[0][0].ToDisplayValue());
});
await Check("SELECT side effects and dynamic SQL are rejected even in nested expressions", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    foreach (var sql in new[] {
        "SELECT * FROM enable_logging()", "SELECT * FROM disable_logging()", "SELECT * FROM enable_profiling()",
        "SELECT * FROM checkpoint()", "SELECT * FROM force_checkpoint()", "SELECT * FROM truncate_duckdb_logs()",
        "SELECT write_log('x')", "SELECT setseed(0.5)", "SELECT nextval('s')",
        "SELECT * FROM query('SELECT 1')", "SELECT * FROM query_table('duckdb_tables')",
        "SELECT * FROM json_execute_serialized_sql('{}')", "SELECT json_serialize_plan('SELECT 1')",
        "SELECT * FROM (SELECT * FROM MAIN.\"enable_logging\"()) q", "SELECT (SELECT write_log('hidden'))",
        "WITH t AS (SELECT write_log('hidden')) SELECT * FROM t", "SELECT list_transform([1], x -> write_log('hidden'))",
        "SELECT * FROM postgres_query('connection', 'DELETE FROM t RETURNING *')", "SELECT * FROM read_duckdb('db.duckdb')",
        "SELECT * FROM histogram_values('duckdb_tables', 'table_name')", "SELECT pg_sleep(10)"
    }) await Reject(async () => await Query(session, sql), "read-and-compute SQL");
});
await Check("replacement scans and unowned internal tables cannot bypass preparation", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    await Reject(async () => await Query(session, $"SELECT * FROM {Literal(csv)}"), "explicit supported reader");
    using var table = session.Acquire(ComputationRequest.Query("SELECT 42 n")); await table.Completion;
    await Reject(async () => await Query(session, "SELECT * FROM " + SqlTemplate.QuoteIdentifier(table.Handle.TableName)), "table handle");
    var name = SqlTemplate.QuoteIdentifier(table.Handle.TableName);
    await Reject(async () => await Query(session, $"WITH {name} AS (SELECT * FROM {name}) SELECT * FROM {name}"), "table handle");
    await Reject(async () => await Query(session, $"WITH RECURSIVE {name}(n) AS (SELECT * FROM {name} UNION ALL SELECT n+1 FROM {name} WHERE n<0) SELECT * FROM {name}"), "table handle");
    var path = SqlTemplate.QuoteIdentifier(csv);
    await Reject(async () => await Query(session, $"WITH {path} AS (SELECT * FROM {path}) SELECT * FROM {path}"), "explicit supported reader");
    Equal(42d, (await Query(session, "SELECT * FROM ?", QueryArgument.Reference(table.Handle))).Rows[0][0].ToDisplayValue());
});
await Check("CSV rejection tables cannot create unmanaged mutable state", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    foreach (var option in new[] { "store_rejects=true", "STORE_REJECTS=true", "store_rejects:=true", "rejects_table='surprise'", "rejects_scan='surprise_scan'" })
        await Reject(async () => await Query(session, $"SELECT * FROM MAIN.READ_CSV(?, {option})", QueryArgument.Scalar(csv)), "reject-table");
    Equal(0d, (await Query(session, "SELECT count(*) FROM duckdb_tables() WHERE table_name IN ('surprise','surprise_scan','reject_errors','reject_scans')")).Rows[0][0].ToDisplayValue());
});
await Check("fragments and plot materialization apply the same capability checks", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    using var bad = session.Acquire(ComputationRequest.Fragment("SELECT * FROM enable_logging()"));
    await Reject(async () => await bad.Completion, "read-and-compute SQL");
    using var plot = session.Acquire(ComputationRequest.Plot(bad.Handle, "bar"));
    await Reject(async () => await plot.Completion, "read-and-compute SQL");
    using var good = session.Acquire(ComputationRequest.Fragment("SELECT * FROM read_csv(?)", QueryArgument.Scalar(csv)));
    await good.Completion;
    using var data = session.AcquireData(good.Handle); await data.Completion;
    Equal(2L, (await session.ReadAsync(data.Handle, new(ReadPurpose.Preview))).TotalRows);
});
await Check("configuration is locked; logging, autoload and persistent secrets are disabled", async () =>
{
    await using var session = Session(new(EnableExternalReads: true));
    foreach (var setting in new[] { "autoinstall_known_extensions", "autoload_known_extensions", "allow_unsigned_extensions", "allow_community_extensions", "allow_unredacted_secrets", "allow_persistent_secrets", "enable_http_logging" })
        Equal("false", (await Query(session, "SELECT current_setting(?)", QueryArgument.Scalar(setting))).Rows[0][0].Text);
    Equal("true", (await Query(session, "SELECT current_setting('lock_configuration')")).Rows[0][0].Text);
    Equal("0", (await Query(session, "SELECT value FROM duckdb_settings() WHERE name='enable_logging'")).Rows[0][0].Text);
});
await Check("native errors omit SQL excerpts and bound string values", async () =>
{
    await using var session = Session();
    const string sensitive = "private-argument-12345";
    var error = await Reject(async () => await Query(session, "SELECT CAST(? AS INTEGER)", QueryArgument.Scalar(sensitive)), "Conversion");
    Equal(false, error.Contains(sensitive)); Equal(false, error.Contains("LINE 1:"));
});
await Check("connection settings reject injection in structural fields", () =>
{
    foreach (var connection in new[] { new S3Connection("not-a-scope"), new S3Connection("s3://bucket/", Region: "x';COPY"), new S3Connection("s3://bucket/", Endpoint: "user:password@host"), new S3Connection("s3://bucket/", S3CredentialSource.AwsProfile) })
    {
        try { connection.Validate(); throw new Exception("Expected invalid connection."); } catch (ArgumentException) { }
    }
    return Task.CompletedTask;
});
if (network)
{
    await using var server = new ReaderServer(File.ReadAllBytes(csv), File.ReadAllBytes(parquet));
    // Set profile locations before the first native AWS SDK initialization. The direct-mode
    // SDK caches these paths for the process; real workbook workers each get a fresh process.
    var credentialsFile = Path.Combine(files.Path, "aws-credentials");
    var configFile = Path.Combine(files.Path, "aws-config");
    File.WriteAllText(credentialsFile, $"[xlduck-fixture]\naws_access_key_id={ReaderServer.AccessKey}\naws_secret_access_key=fixture-secret-never-a-real-key\naws_session_token={ReaderServer.SessionToken}\n");
    File.WriteAllText(configFile, "[profile xlduck-fixture]\nregion=us-east-1\n");
    var oldCredentials = Environment.GetEnvironmentVariable("AWS_SHARED_CREDENTIALS_FILE");
    var oldConfig = Environment.GetEnvironmentVariable("AWS_CONFIG_FILE");
    try
    {
        Environment.SetEnvironmentVariable("AWS_SHARED_CREDENTIALS_FILE", credentialsFile);
        Environment.SetEnvironmentVariable("AWS_CONFIG_FILE", configFile);
        await Check("HTTP and signed S3 reads issue only GET/HEAD; COPY issues no request", async () =>
        {
            var names = new[] { "AWS_ACCESS_KEY_ID", "AWS_SECRET_ACCESS_KEY", "AWS_SESSION_TOKEN" };
            var before = names.Select(Environment.GetEnvironmentVariable).ToArray();
            try
            {
                // Dummy credentials exist only in this test process and its private child.
                Environment.SetEnvironmentVariable(names[0], ReaderServer.AccessKey);
                Environment.SetEnvironmentVariable(names[1], "fixture-secret-never-a-real-key");
                Environment.SetEnvironmentVariable(names[2], ReaderServer.SessionToken);
                var config = new ConnectionSettings(true, extensionPath, [new("s3://bucket/", Endpoint: server.Endpoint, UseSsl: false, PathStyle: true)]);
                await using var session = Session(config);
                Equal(2L, (await Query(session, "SELECT * FROM read_csv(?)", QueryArgument.Scalar(server.BaseUrl + "sample.csv"))).TotalRows);
                Equal(2L, (await Query(session, "SELECT * FROM read_parquet(?)", QueryArgument.Scalar("s3://bucket/sample.parquet"))).TotalRows);
                Equal(2L, (await Query(session, "SELECT * FROM read_parquet(?)", QueryArgument.Scalar("s3://bucket/*.parquet"))).TotalRows);
                Equal(true, server.AuthenticatedS3Reads > 0);
                Equal(0d, (await Query(session, "SELECT count(*) FROM duckdb_settings() WHERE name IN ('s3_access_key_id','s3_secret_access_key','s3_session_token') AND coalesce(value,'')<>''")).Rows[0][0].ToDisplayValue());
                var requestsBefore = server.Requests.Count;
                await Reject(async () => await Query(session, "COPY (SELECT 1) TO 's3://bucket/forbidden.parquet'"), "read-and-compute SELECT");
                Equal(requestsBefore, server.Requests.Count);
                Equal(true, server.Requests.All(method => method is "GET" or "HEAD"));
            }
            finally { for (var i = 0; i < names.Length; i++) Environment.SetEnvironmentVariable(names[i], before[i]); }
        });
        await Check("explicit access keys read compatible S3 endpoints without AWS profiles", async () =>
        {
            var id = Guid.NewGuid().ToString("N");
            var keys = System.Collections.Immutable.ImmutableDictionary<string, S3AccessKey>.Empty.Add(id,
                new(ReaderServer.AccessKey, "fixture-secret-never-a-real-key", ReaderServer.SessionToken));
            var config = new ConnectionSettings(true, extensionPath,
                [new("s3://bucket/", S3CredentialSource.AccessKey, Endpoint: server.Endpoint, UseSsl: false, PathStyle: true, CredentialId: id)], keys);
            var before = server.AuthenticatedS3Reads;
            await using var session = Session(config);
            Equal(2L, (await Query(session, "SELECT * FROM read_parquet(?)", QueryArgument.Scalar("s3://bucket/sample.parquet"))).TotalRows);
            Equal(true, server.AuthenticatedS3Reads > before);
            Equal(0d, (await Query(session, "SELECT count(*) FROM duckdb_settings() WHERE name IN ('s3_access_key_id','s3_secret_access_key','s3_session_token') AND coalesce(value,'')<>''")).Rows[0][0].ToDisplayValue());
        });
        await Check("S3 outage preserves retained data and failed reads recover only on explicit retry", async () =>
        {
            var id = Guid.NewGuid().ToString("N");
            var settings = new ConnectionSettings(true, extensionPath,
                [new("s3://bucket/", S3CredentialSource.AccessKey, Endpoint: server.Endpoint, UseSsl: false, PathStyle: true, CredentialId: id)],
                System.Collections.Immutable.ImmutableDictionary<string, S3AccessKey>.Empty.Add(id,
                    new(ReaderServer.AccessKey, "fixture-secret-never-a-real-key", ReaderServer.SessionToken)));
            await using var session = Session(settings);
            var request = ComputationRequest.Query("SELECT * FROM read_parquet(?)", QueryArgument.Scalar("s3://bucket/sample.parquet"));
            using var retained = session.Acquire(request); await retained.Completion;
            var newRead = ComputationRequest.Query("SELECT * FROM read_csv(?)", QueryArgument.Scalar("s3://bucket/sample.csv"));
            try
            {
                server.Unavailable = true;
                var before = server.Requests.Count;
                using var reused = session.Acquire(request); await reused.Completion;
                Equal(retained.Handle, reused.Handle);
                Equal(2L, (await session.ReadAsync(reused.Handle, new(ReadPurpose.Output))).TotalRows);
                Equal(before, server.Requests.Count);
                using var failedRead = session.Acquire(newRead);
                await Reject(async () => await failedRead.Completion, "503");
                Equal(true, server.Requests.Count > before);
                server.Unavailable = false;
                before = server.Requests.Count;
                using var stillFailed = session.Acquire(newRead);
                Equal(failedRead.Handle, stillFailed.Handle);
                await Reject(async () => await stillFailed.Completion, "503");
                Equal(before, server.Requests.Count);
                session.InvalidateFailuresForRetry();
                using var retried = session.Acquire(newRead); await retried.Completion;
                Equal(false, failedRead.Handle == retried.Handle);
                Equal(2L, (await session.ReadAsync(retried.Handle, new(ReadPurpose.Output))).TotalRows);
                Equal(2L, (await session.ReadAsync(retained.Handle, new(ReadPurpose.Output))).TotalRows);
                Equal(3d, (await Query(session, "SELECT 1+2")).Rows[0][0].ToDisplayValue());
            }
            finally { server.Unavailable = false; }
        });
        await Check("HTTP failures redact signed URL parameters", async () =>
        {
            await using var session = Session(new(true, extensionPath));
            var requestsBefore = server.Requests.Count;
            var error = await Reject(async () => await Query(session, "SELECT * FROM read_csv(?)", QueryArgument.Scalar(server.BaseUrl + "missing.csv?X-Amz-Signature=fixture-signature")), "404");
            Equal(true, server.Requests.Count > requestsBefore);
            Equal(false, error.Contains("fixture-signature"));
        });
        await Check("missing reader installation fails without downloading", async () =>
        {
            await using var session = Session(new(true, files.Path));
            await Reject(async () => await Query(session, "SELECT 1"), "reader extensions are missing");
        });
        await Check("a named AWS profile loads only the configured fixture credentials", async () =>
        {
            var settings = new ConnectionSettings(true, extensionPath, [new("s3://bucket/", S3CredentialSource.AwsProfile, Endpoint: server.Endpoint, Profile: "xlduck-fixture", UseSsl: false, PathStyle: true)]);
            await using var session = Session(settings);
            Equal(2L, (await Query(session, "SELECT * FROM read_csv(?)", QueryArgument.Scalar("s3://bucket/sample.csv"))).TotalRows);
            Equal(0d, (await Query(session, "SELECT count(*) FROM duckdb_settings() WHERE name IN ('s3_access_key_id','s3_secret_access_key','s3_session_token') AND coalesce(value,'')<>''")).Rows[0][0].ToDisplayValue());
            await using var invalid = Session(settings with { S3 = [settings.S3[0] with { Profile = "xlduck-fixture-missing" }] });
            var error = await Reject(async () => await Query(invalid, "SELECT 1"), "could not load its credentials");
            Equal(false, error.Contains("fixture-secret-never-a-real-key"));
        });
        await Check("modified reader binaries are rejected before loading", async () =>
        {
            File.WriteAllText(Path.Combine(files.Path, "httpfs.duckdb_extension"), "not the pinned extension");
            await using var session = Session(new(true, files.Path));
            await Reject(async () => await Query(session, "SELECT 1"), "does not match the pinned");
        });
    }
    finally { Environment.SetEnvironmentVariable("AWS_SHARED_CREDENTIALS_FILE", oldCredentials); Environment.SetEnvironmentVariable("AWS_CONFIG_FILE", oldConfig); }
}
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
