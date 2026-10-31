// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Diagnostics;
using System.Data;
using DuckDB.NET.Data;

namespace XlDuck.ExecutionPrototype;

public sealed class LocalBackend : IBackend
{
    private readonly SemaphoreSlim _gate = new(1);
    private readonly Dictionary<string, string> _tables = new();
    private readonly CancellationTokenSource _lifetime = new();
    private readonly string _session = Guid.NewGuid().ToString("N");
    private DuckDBConnection? _connection;
    private int _nextId;

    public async Task<Result> ExecuteAsync(Request request, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetime.Token);
        await _gate.WaitAsync(linked.Token).ConfigureAwait(false);
        try
        {
            return await Task.Run(() => Execute(request, linked.Token), linked.Token).ConfigureAwait(false);
        }
        finally { _gate.Release(); }
    }

    private Result Execute(Request request, CancellationToken token)
    {
        if (_connection is null)
        {
            _connection = new DuckDBConnection("DataSource=:memory:");
            _connection.Open();
            NonQuery("SET threads=2; SET memory_limit='256MB'; SET enable_external_access=false;");
        }
        token.ThrowIfCancellationRequested();
        var watch = Stopwatch.StartNew();
        using var interruption = token.Register(() => _connection.NativeConnection.Interrupt());
        string? partialTable = null;
        try
        {
            if (request.Operation == "rows")
            {
                var table = Resolve(request.Input);
                using var command = _connection.CreateCommand();
                command.CommandText = $"SELECT * FROM \"{table}\" LIMIT {Math.Clamp(request.Limit, 1, 10000)}";
                using var reader = command.ExecuteReader();
                var columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
                var rows = new List<Atom[]>();
                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();
                    rows.Add(Enumerable.Range(0, reader.FieldCount).Select(i => Atom.From(reader.GetValue(i))).ToArray());
                }
                return new(request.Id, request.Input, columns, rows.ToArray(), rows.Count, watch.Elapsed.TotalMilliseconds);
            }
            if (request.Operation == "capture")
            {
                partialTable = "result_" + ++_nextId;
                var rows = request.Capture ?? throw new ArgumentException("Capture data missing.");
                NonQuery($"CREATE TEMP TABLE \"{partialTable}\" (value DOUBLE)");
                using var appender = _connection.CreateAppender(partialTable);
                foreach (var row in rows)
                {
                    token.ThrowIfCancellationRequested();
                    appender.CreateRow().AppendValue(Convert.ToDouble(row[0].ToObject())).EndRow();
                }
                appender.Close();
            }
            else if (request.Operation == "query")
            {
                if (request.Input.Length > 0)
                    NonQuery($"CREATE OR REPLACE TEMP VIEW __input AS SELECT * FROM \"{Resolve(request.Input)}\"");
                else
                    NonQuery("DROP VIEW IF EXISTS __input");
                token.ThrowIfCancellationRequested();
                partialTable = "result_" + ++_nextId;
                using var command = _connection.CreateCommand();
                command.CommandText = $"CREATE TEMP TABLE \"{partialTable}\" AS {request.Sql}";
                if (request.Value is not null)
                {
                    var type = request.Value.Kind switch
                    {
                        "number" => DbType.Double,
                        "boolean" => DbType.Boolean,
                        _ => DbType.String
                    };
                    command.Parameters.Add(new DuckDBParameter("value", type, request.Value.ToObject()));
                }
                command.ExecuteNonQuery();
            }
            else { throw new ArgumentException($"Unknown operation {request.Operation}."); }

            token.ThrowIfCancellationRequested();
            using var count = _connection.CreateCommand();
            count.CommandText = $"SELECT COUNT(*) FROM \"{partialTable}\"";
            var rowCount = Convert.ToInt64(count.ExecuteScalar());
            var handle = $"proto://{_session}/{_nextId}";
            _tables.Add(handle, partialTable);
            partialTable = null;
            return new(request.Id, Handle: handle, RowCount: rowCount, EngineMs: watch.Elapsed.TotalMilliseconds);
        }
        catch when (token.IsCancellationRequested) { throw new OperationCanceledException(token); }
        finally
        {
            if (partialTable is not null)
            {
                try { NonQuery($"DROP TABLE IF EXISTS \"{partialTable}\""); }
                catch { /* Session disposal remains the final cleanup boundary. */ }
            }
        }
    }

    private string Resolve(string handle) => _tables.TryGetValue(handle, out var table)
        ? table : throw new ArgumentException("Handle does not belong to this session.");
    private void NonQuery(string sql)
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }
    public async ValueTask DisposeAsync()
    {
        _lifetime.Cancel();
        await _gate.WaitAsync().ConfigureAwait(false);
        try { _connection?.Dispose(); _tables.Clear(); }
        finally { _gate.Release(); }
    }
}
