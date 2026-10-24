// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Data;
using System.Globalization;
using System.Numerics;
using System.Text.Json;
using DuckDB.NET.Data;
using XlDuck.Core;

namespace XlDuck.Engine;

/// <summary>Serial, session-owned storage. External access stays disabled until the reader capability policy is implemented.</summary>
public sealed class DuckDbBackend(Guid sessionId, string? spillDirectory = null) : IExecutionBackend
{
    private readonly SessionDirectory? _ownedDirectory = spillDirectory is null ? new() : null;
    private readonly SemaphoreSlim _gate = new(1);
    private readonly CancellationTokenSource _lifetime = new();
    private readonly Dictionary<ResultHandle, ComputationResult> _tables = new();
    private DuckDBConnection? _connection;
    private Task? _disposal;
    private readonly object _disposeSync = new();
    public int TableCount { get; private set; }

    private async Task<T> WithConnection<T>(Func<CancellationToken, T> action, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetime.Token);
        await _gate.WaitAsync(linked.Token).ConfigureAwait(false);
        try
        {
            return await Task.Run(() =>
            {
                linked.Token.ThrowIfCancellationRequested();
                if (_connection is null)
                {
                    var connection = new DuckDBConnection("DataSource=:memory:");
                    try
                    {
                        connection.Open();
                        using var setup = connection.CreateCommand();
                        var temporaryPath = spillDirectory ?? _ownedDirectory!.Path;
                        setup.CommandText = $"SET threads=2; SET memory_limit='512MB'; SET temp_directory='{temporaryPath.Replace("'", "''", StringComparison.Ordinal)}'; SET max_temp_directory_size='1GiB'; SET autoinstall_known_extensions=false; SET autoload_known_extensions=false; SET enable_external_access=false;";
                        setup.ExecuteNonQuery();
                        _connection = connection;
                    }
                    catch { connection.Dispose(); throw; }
                }
                using var interrupt = linked.Token.Register(() => _connection.NativeConnection.Interrupt());
                try { return action(linked.Token); }
                catch when (linked.IsCancellationRequested) { throw new OperationCanceledException(linked.Token); }
            }, linked.Token).ConfigureAwait(false);
        }
        finally { _gate.Release(); }
    }

    public Task<ComputationResult> ExecuteAsync(BackendRequest request, CancellationToken token) => WithConnection(ct =>
    {
        ValidateHandle(request.Output);
        if (_tables.ContainsKey(request.Output)) throw new ArgumentException("Output handle already exists.");
        var name = SqlTemplate.QuoteIdentifier(request.Output.TableName);
        var committed = false;
        try
        {
            if (request.Capture is { } capture) Capture(request.Output.TableName, capture, ct);
            else
            {
                var query = request.Query ?? throw new ArgumentException("Query is missing.");
                foreach (var input in query.Tables) Resolve(input);
                ValidateSelect(query.Sql);
                using var command = _connection!.CreateCommand();
                command.CommandText = $"CREATE TEMP TABLE {name} AS {query.Sql}";
                foreach (var parameter in query.Parameters) AddParameter(command, parameter.Name, parameter.Value);
                command.ExecuteNonQuery();
            }
            ct.ThrowIfCancellationRequested();
            using var schema = _connection!.CreateCommand();
            schema.CommandText = $"DESCRIBE SELECT * FROM {name}";
            var columns = ImmutableArray.CreateBuilder<ColumnSchema>();
            using (var reader = schema.ExecuteReader())
                while (reader.Read()) columns.Add(new(reader.GetString(0), reader.GetString(1)));
            using var count = _connection.CreateCommand();
            count.CommandText = $"SELECT COUNT(*) FROM {name}";
            var result = new ComputationResult(request.Output, columns.ToImmutable(), Convert.ToInt64(count.ExecuteScalar(), CultureInfo.InvariantCulture));
            ct.ThrowIfCancellationRequested();
            _tables.Add(result.Handle, result);
            TableCount = _tables.Count;
            committed = true;
            return result;
        }
        finally
        {
            if (!committed) NonQuery($"DROP TABLE IF EXISTS {name}");
        }
    }, token);

    // Use DuckDB's parser before execution, rather than a SELECT-prefix filter. This is not the full extension/function policy.
    private void ValidateSelect(string sql)
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = "SELECT json_serialize_sql(CAST($sql AS VARCHAR))";
        command.Parameters.Add(new DuckDBParameter("sql", DbType.String, sql));
        using var parsed = JsonDocument.Parse((string)command.ExecuteScalar()!);
        var root = parsed.RootElement;
        if (root.GetProperty("error").GetBoolean() || root.GetProperty("statements").GetArrayLength() != 1)
            throw new ArgumentException("Only one read-and-compute SELECT query is supported. DDL, DML, settings and multiple statements are not allowed.");
    }

    public Task<TablePage> ReadAsync(ResultHandle handle, ReadRequest request, CancellationToken token) => WithConnection(ct =>
    {
        var table = Resolve(handle);
        if (request.MaxRows < 1 || request.MaxCells < 1) throw new ArgumentOutOfRangeException(nameof(request));
        if (request.Purpose == ReadPurpose.Scalar && (table.RowCount != 1 || table.Columns.Length != 1))
            throw new InvalidOperationException("Scalar output requires exactly one row and one column. Aggregate or filter the query explicitly.");
        var maxRows = Math.Min(request.MaxRows, Math.Max(0, request.MaxCells / Math.Max(1, table.Columns.Length) - 1));
        if (maxRows < 1) throw new InvalidOperationException("The column count exceeds the output cell limit.");
        if (request.Purpose == ReadPurpose.Output && table.RowCount > maxRows)
            throw new InvalidOperationException($"Result has {table.RowCount} rows and exceeds the output limit of {maxRows}. Filter or aggregate it, or request a sampled preview.");
        if (request.Purpose == ReadPurpose.Scalar) maxRows = 1;
        var projections = table.Columns.Select(column => IsNested(column.Type)
            ? $"to_json({SqlTemplate.QuoteIdentifier(column.Name)})::VARCHAR"
            : $"CAST({SqlTemplate.QuoteIdentifier(column.Name)} AS VARCHAR)");
        using var command = _connection!.CreateCommand();
        command.CommandText = $"SELECT {string.Join(',', projections)} FROM {SqlTemplate.QuoteIdentifier(handle.TableName)} LIMIT {maxRows}";
        var rows = ImmutableArray.CreateBuilder<ImmutableArray<CellValue>>();
        long characters = 0;
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            ct.ThrowIfCancellationRequested();
            var row = ImmutableArray.CreateBuilder<CellValue>();
            for (var column = 0; column < table.Columns.Length; column++)
            {
                var value = reader.IsDBNull(column) ? CellValue.Null : ParseValue(table.Columns[column].Type, reader.GetString(column));
                characters += value.Text.Length;
                if (characters > 8 * 1024 * 1024) throw new InvalidOperationException("Output text exceeds 8 MiB of characters. Filter or shorten the requested output.");
                row.Add(value);
            }
            rows.Add(row.ToImmutable());
        }
        return new TablePage(table.Columns, rows.ToImmutable(), table.RowCount, table.RowCount > rows.Count);
    }, token);

    private static bool IsNested(string type) => type.EndsWith(']') || type.StartsWith("STRUCT(", StringComparison.Ordinal) ||
        type.StartsWith("MAP(", StringComparison.Ordinal) || type.StartsWith("UNION(", StringComparison.Ordinal) || type == "JSON";
    private static CellValue ParseValue(string type, string text)
    {
        if (IsNested(type)) return new(ValueKind.Json, text);
        if (type.StartsWith("DECIMAL(", StringComparison.Ordinal))
            return new(ValueKind.Decimal, text.StartsWith('.') ? "0" + text : text.StartsWith("-.", StringComparison.Ordinal) ? "-0" + text[1..] : text);
        return type switch
        {
            "TINYINT" or "SMALLINT" or "INTEGER" or "BIGINT" or "HUGEINT" or "UTINYINT" or "USMALLINT" or "UINTEGER" or "UBIGINT" or "UHUGEINT" or "BIGNUM" => new(ValueKind.Integer, text),
            "FLOAT" or "DOUBLE" => CellValue.Number(double.Parse(text, CultureInfo.InvariantCulture)),
            "BOOLEAN" => new(ValueKind.Boolean, text),
            "DATE" => new(ValueKind.Date, text),
            "TIME" => new(ValueKind.Time, text),
            "TIMESTAMP" or "TIMESTAMP_S" or "TIMESTAMP_MS" => new(ValueKind.Timestamp, text),
            "TIMESTAMP WITH TIME ZONE" => new(ValueKind.TimestampOffset, text),
            "INTERVAL" => new(ValueKind.Interval, text),
            // Nanosecond timestamps, time zones, blobs, UUIDs and extension types stay exact display text.
            _ => CellValue.String(text)
        };
    }

    private void Capture(string tableName, CaptureData capture, CancellationToken token)
    {
        capture.Validate();
        var headers = CaptureData.UniqueHeaders(capture.Headers);
        var types = Enumerable.Range(0, headers.Length).Select(column => CaptureType(capture, column)).ToArray();
        NonQuery($"CREATE TEMP TABLE {SqlTemplate.QuoteIdentifier(tableName)} ({string.Join(',', headers.Select((header, i) => SqlTemplate.QuoteIdentifier(header) + " " + types[i]))})");
        using var appender = _connection!.CreateAppender(tableName);
        foreach (var values in capture.Rows)
        {
            token.ThrowIfCancellationRequested();
            appender.AppendRow(values, static (row, cells) =>
            {
                foreach (var value in cells)
                {
                    // Match the physical column type exactly; an appender is not an implicit conversion API.
                    switch (value.Kind)
                    {
                        case ValueKind.Null or ValueKind.Empty: row.AppendNullValue(); break;
                        case ValueKind.Text or ValueKind.Json: row.AppendValue(value.Text); break;
                        case ValueKind.Number: row.AppendValue(double.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        case ValueKind.Integer: row.AppendValue(BigInteger.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        case ValueKind.Boolean: row.AppendValue(bool.Parse(value.Text)); break;
                        case ValueKind.Date: row.AppendValue(DateOnly.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        case ValueKind.Time: row.AppendValue(TimeOnly.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        case ValueKind.Timestamp: row.AppendValue(DateTime.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        case ValueKind.TimestampOffset: row.AppendValue(DateTimeOffset.Parse(value.Text, CultureInfo.InvariantCulture)); break;
                        default: throw new ArgumentException($"Unsupported capture value {value.Kind}.");
                    }
                }
            });
        }
        appender.Close();
    }
    private static string CaptureType(CaptureData capture, int column)
    {
        var kinds = capture.Rows.Select(row => row[column].Kind).Where(kind => kind is not (ValueKind.Empty or ValueKind.Null)).Distinct().ToArray();
        if (kinds.Length > 1) throw new ArgumentException($"Capture column {column + 1} mixes value types. Normalize the column explicitly before capture.");
        return kinds.SingleOrDefault() switch
        {
            ValueKind.Null => "VARCHAR", ValueKind.Text => "VARCHAR", ValueKind.Boolean => "BOOLEAN", ValueKind.Number => "DOUBLE",
            ValueKind.Integer => "HUGEINT", ValueKind.Date => "DATE", ValueKind.Time => "TIME", ValueKind.Timestamp => "TIMESTAMP",
            ValueKind.TimestampOffset => "TIMESTAMPTZ", ValueKind.Json => "JSON",
            _ => throw new ArgumentException($"Capture column {column + 1} has an unsupported type. Convert it explicitly before capture.")
        };
    }
    private static void AddParameter(DuckDBCommand command, string name, CellValue value)
    {
        value.Validate();
        (DbType Type, object Value) parameter = value.Kind switch
        {
            ValueKind.Empty or ValueKind.Null => (DbType.String, DBNull.Value),
            ValueKind.Text => (DbType.String, value.Text),
            ValueKind.Boolean => (DbType.Boolean, bool.Parse(value.Text)),
            ValueKind.Number => (DbType.Double, double.Parse(value.Text, CultureInfo.InvariantCulture)),
            ValueKind.Integer when long.TryParse(value.Text, CultureInfo.InvariantCulture, out var integer) => (DbType.Int64, integer),
            ValueKind.Integer => (DbType.VarNumeric, BigInteger.Parse(value.Text, CultureInfo.InvariantCulture)),
            // SQL preparation supplies the exact DECIMAL cast; .NET decimal would round values above 29 digits.
            ValueKind.Decimal => (DbType.String, value.Text),
            ValueKind.Date => (DbType.Date, DateOnly.Parse(value.Text, CultureInfo.InvariantCulture)),
            ValueKind.Time => (DbType.Time, TimeOnly.Parse(value.Text, CultureInfo.InvariantCulture)),
            ValueKind.Timestamp => (DbType.DateTime2, DateTime.Parse(value.Text, CultureInfo.InvariantCulture)),
            ValueKind.TimestampOffset => (DbType.DateTimeOffset, DateTimeOffset.Parse(value.Text, CultureInfo.InvariantCulture)),
            ValueKind.Json => (DbType.String, value.Text),
            _ => throw new ArgumentException($"Unsupported scalar argument type {value.Kind}. Convert it explicitly in SQL.")
        };
        command.Parameters.Add(new DuckDBParameter(name, parameter.Type, parameter.Value));
    }

    public async Task DropAsync(ResultHandle handle) => await WithConnection(_ =>
    {
        ValidateHandle(handle);
        if (_tables.Remove(handle)) NonQuery($"DROP TABLE IF EXISTS {SqlTemplate.QuoteIdentifier(handle.TableName)}");
        TableCount = _tables.Count;
        return true;
    }, CancellationToken.None).ConfigureAwait(false);
    private ComputationResult Resolve(ResultHandle handle)
    {
        ValidateHandle(handle);
        return _tables.GetValueOrDefault(handle) ?? throw new ArgumentException("Table is not retained in this session.");
    }
    private void ValidateHandle(ResultHandle handle)
    {
        if (handle.Session != sessionId || handle.Kind != ResultKind.Table) throw new ArgumentException("Table handle does not belong to this engine session.");
    }
    private void NonQuery(string sql)
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }
    public ValueTask DisposeAsync()
    {
        lock (_disposeSync) return new(_disposal ??= DisposeCoreAsync());
    }
    private async Task DisposeCoreAsync()
    {
        _lifetime.Cancel();
        await _gate.WaitAsync().ConfigureAwait(false);
        try { _connection?.Dispose(); _connection = null; _tables.Clear(); TableCount = 0; _ownedDirectory?.Dispose(); }
        finally { _gate.Release(); }
    }
}
