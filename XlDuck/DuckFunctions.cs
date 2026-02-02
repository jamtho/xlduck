// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Text.RegularExpressions;
using ExcelDna.Integration;
using DuckDB.NET.Data;

namespace XlDuck;

/// <summary>
/// Add-in lifecycle handler.
/// </summary>
public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        System.Diagnostics.Debug.WriteLine("[XlDuck] Add-in loaded");
    }

    public void AutoClose()
    {
        System.Diagnostics.Debug.WriteLine("[XlDuck] Add-in unloaded");
    }
}

public static class DuckFunctions
{
    private const string Version = "0.1";

    private static DuckDBConnection? _connection;
    private static readonly object _connLock = new();

    // Ready flag - queries with "@config" sentinel wait until DuckConfigReady() is called
    internal static bool IsReady { get; private set; } = false;

    // Sentinel value that marks a query as requiring config
    internal const string ConfigSentinel = "@config";

    // Status URL prefixes (# prefix follows Excel convention)
    internal const string BlockedPrefix = "#duck://blocked/";
    internal const string ErrorPrefix = "#duck://error/";
    internal const string ConfigBlockedStatus = "#duck://blocked/config|Waiting for DuckConfigReady()";

    /// <summary>
    /// Format an error message as a duck:// URL.
    /// </summary>
    internal static string FormatError(string category, string message)
    {
        // Truncate long messages and remove newlines
        var cleanMessage = message.Replace("\r", "").Replace("\n", " ");
        if (cleanMessage.Length > 200)
            cleanMessage = cleanMessage.Substring(0, 197) + "...";
        return $"{ErrorPrefix}{category}|{cleanMessage}";
    }

    /// <summary>
    /// Format an exception as a duck:// error URL with auto-categorization.
    /// </summary>
    internal static string FormatException(Exception ex)
    {
        var msg = ex.Message;
        string category;

        if (msg.Contains("Parser Error") || msg.Contains("syntax error"))
            category = "syntax";
        else if (msg.Contains("does not exist") || msg.Contains("not found"))
            category = "notfound";
        else if (msg.Contains("HTTP"))
            category = "http";
        else
            category = "query";

        return FormatError(category, msg);
    }

    /// <summary>
    /// Check if a value is a duck:// error or blocked status.
    /// </summary>
    internal static bool IsErrorOrBlocked(string? value)
    {
        return value != null && (value.StartsWith(ErrorPrefix) || value.StartsWith(BlockedPrefix));
    }

    private static DuckDBConnection GetConnection()
    {
        if (_connection == null)
        {
            lock (_connLock)
            {
                _connection ??= new DuckDBConnection("DataSource=:memory:");
                _connection.Open();
            }
        }
        return _connection;
    }

    [ExcelFunction(Description = "Get the XlDuck add-in version")]
    public static string DuckVersion()
    {
        return Version;
    }

    [ExcelFunction(Description = "Signal that configuration is complete. Queries with @config wait until this is called.")]
    public static string DuckConfigReady()
    {
        System.Diagnostics.Debug.WriteLine("[XlDuck] DuckConfigReady called");
        IsReady = true;
        return "OK";
    }

    [ExcelFunction(Description = "Get the DuckDB library version")]
    public static string DuckLibraryVersion()
    {
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT version()";
            return cmd.ExecuteScalar()?.ToString() ?? "Unknown";
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL query and return a handle. Use :name placeholders with name/value pairs.")]
    public static object DuckQuery(
        [ExcelArgument(Description = "SQL query with optional :name placeholders")] string sql,
        [ExcelArgument(Description = "First parameter name")] object name1 = null!,
        [ExcelArgument(Description = "First parameter value")] object value1 = null!,
        [ExcelArgument(Description = "Second parameter name")] object name2 = null!,
        [ExcelArgument(Description = "Second parameter value")] object value2 = null!,
        [ExcelArgument(Description = "Third parameter name")] object name3 = null!,
        [ExcelArgument(Description = "Third parameter value")] object value3 = null!,
        [ExcelArgument(Description = "Fourth parameter name")] object name4 = null!,
        [ExcelArgument(Description = "Fourth parameter value")] object value4 = null!)
    {
        try
        {
            var args = CollectArgs(name1, value1, name2, value2, name3, value3, name4, value4);
            // Build topic info: ["query", sql, arg1, arg2, ...]
            var topicInfo = new List<string> { "query", sql };
            topicInfo.AddRange(args.Select(a => a?.ToString() ?? ""));

            return XlCall.RTD("XlDuck.DuckRtdServer", null, topicInfo.ToArray());
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    [ExcelFunction(Description = "Output a handle as a spilled array with headers.")]
    public static object[,] DuckOut(
        [ExcelArgument(Description = "Handle from DuckQuery or DuckFrag (e.g. duck://table/1 or duck://frag/1)")] string handle)
    {
        try
        {
            if (ResultStore.IsHandle(handle))
            {
                var stored = ResultStore.Get(handle);
                if (stored == null)
                {
                    return new object[,] { { FormatError("notfound", $"Handle not found: {handle}") } };
                }
                return StoredResultToArray(stored);
            }
            else if (FragmentStore.IsHandle(handle))
            {
                // Execute the fragment and output results
                var fragment = FragmentStore.Get(handle);
                if (fragment == null)
                {
                    return new object[,] { { FormatError("notfound", $"Fragment not found: {handle}") } };
                }

                var (resolvedSql, tempTables) = ResolveParameters(fragment.Sql, fragment.Args, new HashSet<string> { handle });
                try
                {
                    var conn = GetConnection();
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = resolvedSql;
                    using var reader = cmd.ExecuteReader();

                    var fieldCount = reader.FieldCount;
                    var columnNames = new string[fieldCount];
                    var columnTypes = new Type[fieldCount];

                    for (int i = 0; i < fieldCount; i++)
                    {
                        columnNames[i] = reader.GetName(i);
                        columnTypes[i] = reader.GetFieldType(i);
                    }

                    var rows = new List<object?[]>();
                    while (reader.Read())
                    {
                        var row = new object?[fieldCount];
                        for (int i = 0; i < fieldCount; i++)
                        {
                            row[i] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        }
                        rows.Add(row);
                    }

                    var stored = new StoredResult(columnNames, columnTypes, rows);
                    return StoredResultToArray(stored);
                }
                finally
                {
                    CleanupTempTables(tempTables);
                }
            }
            else
            {
                return new object[,] { { FormatError("invalid", $"Invalid handle format: {handle}") } };
            }
        }
        catch (Exception ex)
        {
            return new object[,] { { $"#ERROR: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL query and output results as a spilled array. Use :name placeholders with name/value pairs.")]
    public static object[,] DuckQueryOut(
        [ExcelArgument(Description = "SQL query with optional :name placeholders")] string sql,
        [ExcelArgument(Description = "First parameter name")] object name1 = null!,
        [ExcelArgument(Description = "First parameter value")] object value1 = null!,
        [ExcelArgument(Description = "Second parameter name")] object name2 = null!,
        [ExcelArgument(Description = "Second parameter value")] object value2 = null!,
        [ExcelArgument(Description = "Third parameter name")] object name3 = null!,
        [ExcelArgument(Description = "Third parameter value")] object value3 = null!,
        [ExcelArgument(Description = "Fourth parameter name")] object name4 = null!,
        [ExcelArgument(Description = "Fourth parameter value")] object value4 = null!)
    {
        try
        {
            var args = CollectArgs(name1, value1, name2, value2, name3, value3, name4, value4);
            var handle = ExecuteQueryInternal(sql, args);

            if (IsErrorOrBlocked(handle))
            {
                return new object[,] { { handle } };
            }

            var stored = ResultStore.Get(handle);
            if (stored == null)
            {
                return new object[,] { { FormatError("notfound", $"Handle not found: {handle}") } };
            }
            return StoredResultToArray(stored);
        }
        catch (Exception ex)
        {
            return new object[,] { { $"#ERROR: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL statement (CREATE, INSERT, etc.)")]
    public static object DuckExecute(
        [ExcelArgument(Description = "SQL statement")] string sql)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            var conn = GetConnection();
            var connTime = sw.ElapsedMilliseconds;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            var rowsAffected = cmd.ExecuteNonQuery();
            System.Diagnostics.Debug.WriteLine($"[DuckExecute] conn={connTime}ms exec={sw.ElapsedMilliseconds - connTime}ms sql={sql.Substring(0, Math.Min(50, sql.Length))}");
            return $"OK ({rowsAffected} rows affected)";
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    [ExcelFunction(Description = "Create a SQL fragment for lazy evaluation. Use :name placeholders with name/value pairs.")]
    public static object DuckFrag(
        [ExcelArgument(Description = "SQL query (SELECT or PIVOT) with optional :name placeholders")] string sql,
        [ExcelArgument(Description = "First parameter name")] object name1 = null!,
        [ExcelArgument(Description = "First parameter value")] object value1 = null!,
        [ExcelArgument(Description = "Second parameter name")] object name2 = null!,
        [ExcelArgument(Description = "Second parameter value")] object value2 = null!,
        [ExcelArgument(Description = "Third parameter name")] object name3 = null!,
        [ExcelArgument(Description = "Third parameter value")] object value3 = null!,
        [ExcelArgument(Description = "Fourth parameter name")] object name4 = null!,
        [ExcelArgument(Description = "Fourth parameter value")] object value4 = null!)
    {
        try
        {
            var args = CollectArgs(name1, value1, name2, value2, name3, value3, name4, value4);
            // Build topic info: ["frag", sql, arg1, arg2, ...]
            var topicInfo = new List<string> { "frag", sql };
            topicInfo.AddRange(args.Select(a => a?.ToString() ?? ""));

            return XlCall.RTD("XlDuck.DuckRtdServer", null, topicInfo.ToArray());
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    /// <summary>
    /// Execute a query, store the result, and return the handle. Called by RTD server.
    /// </summary>
    internal static string ExecuteQueryInternal(string sql, object[] args)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var (resolvedSql, tempTables) = ResolveParameters(sql, args, new HashSet<string>());
        var resolveTime = sw.ElapsedMilliseconds;

        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = resolvedSql;

            sw.Restart();
            using var reader = cmd.ExecuteReader();
            var executeTime = sw.ElapsedMilliseconds;

            var fieldCount = reader.FieldCount;
            var columnNames = new string[fieldCount];
            var columnTypes = new Type[fieldCount];

            for (int i = 0; i < fieldCount; i++)
            {
                columnNames[i] = reader.GetName(i);
                columnTypes[i] = reader.GetFieldType(i);
            }

            sw.Restart();
            var rows = new List<object?[]>();
            while (reader.Read())
            {
                var row = new object?[fieldCount];
                for (int i = 0; i < fieldCount; i++)
                {
                    row[i] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                }
                rows.Add(row);
            }
            var readTime = sw.ElapsedMilliseconds;

            var storedResult = new StoredResult(columnNames, columnTypes, rows);
            var handle = ResultStore.Store(storedResult);

            System.Diagnostics.Debug.WriteLine($"[DuckQuery] resolve={resolveTime}ms execute={executeTime}ms read={readTime}ms rows={rows.Count}");
            return handle;
        }
        finally
        {
            CleanupTempTables(tempTables);
        }
    }

    /// <summary>
    /// Create a fragment, validate it, and return the handle. Called by RTD server.
    /// </summary>
    internal static string CreateFragmentInternal(string sql, object[] args)
    {
        // Validate the SQL by resolving parameters and running EXPLAIN
        var (resolvedSql, tempTables) = ResolveParameters(sql, args, new HashSet<string>());
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"EXPLAIN {resolvedSql}";
            cmd.ExecuteNonQuery();
        }
        finally
        {
            CleanupTempTables(tempTables);
        }

        // Store the fragment with original SQL and args
        var fragment = new StoredFragment(sql, args);
        return FragmentStore.Store(fragment);
    }

    /// <summary>
    /// Convert a stored result to an Excel array with headers.
    /// </summary>
    private static object[,] StoredResultToArray(StoredResult stored)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var fieldCount = stored.ColumnNames.Length;
        var rowCount = stored.Rows.Count;

        if (fieldCount == 0)
        {
            return new object[,] { { FormatError("query", "No columns") } };
        }

        var result = new object[rowCount + 1, fieldCount];
        var allocTime = sw.ElapsedMilliseconds;

        // Header row
        for (int j = 0; j < fieldCount; j++)
        {
            result[0, j] = stored.ColumnNames[j];
        }

        // Data rows
        sw.Restart();
        for (int i = 0; i < rowCount; i++)
        {
            var row = stored.Rows[i];
            for (int j = 0; j < fieldCount; j++)
            {
                result[i + 1, j] = ConvertToExcelValue(row[j]);
            }
        }
        var copyTime = sw.ElapsedMilliseconds;

        System.Diagnostics.Debug.WriteLine($"[DuckOut] alloc={allocTime}ms copy={copyTime}ms rows={rowCount} cols={fieldCount}");
        return result;
    }

    /// <summary>
    /// Collect optional name/value pairs into an array, filtering out missing values.
    /// </summary>
    private static object[] CollectArgs(params object[] pairs)
    {
        var result = new List<object>();
        for (int i = 0; i < pairs.Length; i += 2)
        {
            var name = pairs[i];
            var value = pairs[i + 1];

            // Skip if name is missing/empty
            if (name == null || name is ExcelMissing || name is ExcelEmpty)
                break;
            if (name is string s && string.IsNullOrEmpty(s))
                break;

            result.Add(name);
            result.Add(value);
        }
        return result.ToArray();
    }

    /// <summary>
    /// Parse SQL for :name placeholders, look up handles, create temp tables or inline fragments, and return resolved SQL.
    /// </summary>
    /// <param name="sql">SQL with :name placeholders</param>
    /// <param name="args">Name/value pairs for parameter binding</param>
    /// <param name="visitedFragments">Set of fragment handles currently being resolved (for cycle detection)</param>
    private static (string resolvedSql, List<string> tempTables) ResolveParameters(string sql, object[] args, HashSet<string> visitedFragments)
    {
        var tempTables = new List<string>();

        if (args.Length == 0)
        {
            return (sql, tempTables);
        }

        if (args.Length % 2 != 0)
        {
            throw new ArgumentException("Parameters must be name/value pairs");
        }

        var parameters = new Dictionary<string, string>();
        for (int i = 0; i < args.Length; i += 2)
        {
            var name = args[i]?.ToString() ?? throw new ArgumentException($"Parameter name at position {i} is null");
            var value = args[i + 1]?.ToString() ?? "";

            if (ResultStore.IsHandle(value))
            {
                // Table handle: create temp table from stored data
                var stored = ResultStore.Get(value) ?? throw new ArgumentException($"Handle not found: {value}");
                var tempTableName = CreateTempTable(stored);
                tempTables.Add(tempTableName);
                parameters[name] = tempTableName;
            }
            else if (FragmentStore.IsHandle(value))
            {
                // Fragment handle: resolve recursively and inline as subquery
                if (visitedFragments.Contains(value))
                {
                    throw new ArgumentException($"Circular fragment reference detected: {value}");
                }

                var fragment = FragmentStore.Get(value) ?? throw new ArgumentException($"Fragment not found: {value}");

                // Add to visited set before recursing
                visitedFragments.Add(value);
                var (resolvedFragmentSql, fragmentTempTables) = ResolveParameters(fragment.Sql, fragment.Args, visitedFragments);
                visitedFragments.Remove(value);

                // Collect any temp tables created during fragment resolution
                tempTables.AddRange(fragmentTempTables);

                // Wrap fragment SQL in parentheses as a subquery
                parameters[name] = $"({resolvedFragmentSql})";
            }
            else
            {
                // Quote string values for SQL (escape single quotes)
                var escaped = value.Replace("'", "''");
                parameters[name] = $"'{escaped}'";
            }
        }

        var resolvedSql = Regex.Replace(sql, @":(\w+)", match =>
        {
            var paramName = match.Groups[1].Value;
            if (parameters.TryGetValue(paramName, out var replacement))
            {
                return replacement;
            }
            return match.Value;
        });

        return (resolvedSql, tempTables);
    }

    /// <summary>
    /// Create a temp table from a stored result and return its name.
    /// </summary>
    private static string CreateTempTable(StoredResult stored)
    {
        var conn = GetConnection();
        var tableName = $"_xlduck_temp_{Guid.NewGuid():N}";

        var columnDefs = new List<string>();
        for (int i = 0; i < stored.ColumnNames.Length; i++)
        {
            var colName = stored.ColumnNames[i];
            var colType = MapTypeToDuckDB(stored.ColumnTypes[i]);
            columnDefs.Add($"\"{colName}\" {colType}");
        }

        var createSql = $"CREATE TEMP TABLE \"{tableName}\" ({string.Join(", ", columnDefs)})";
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = createSql;
            cmd.ExecuteNonQuery();
        }

        if (stored.Rows.Count > 0)
        {
            // Use DuckDB Appender for fast bulk inserts
            using var appender = ((DuckDB.NET.Data.DuckDBConnection)conn).CreateAppender(tableName);
            foreach (var row in stored.Rows)
            {
                var appenderRow = appender.CreateRow();
                foreach (var value in row)
                {
                    AppendTypedValue(appenderRow, value);
                }
                appenderRow.EndRow();
            }
        }

        return tableName;
    }

    /// <summary>
    /// Convert DuckDB values to Excel-compatible types.
    /// Handles HUGEINT, DECIMAL, and other types that Excel/COM doesn't support natively.
    /// </summary>
    private static object ConvertToExcelValue(object? value)
    {
        if (value == null || value == DBNull.Value)
            return "";  // Empty string displays as blank in spilled arrays

        // Handle BigInteger (used for HUGEINT)
        if (value is System.Numerics.BigInteger bigInt)
        {
            if (bigInt >= long.MinValue && bigInt <= long.MaxValue)
                return (double)(long)bigInt;
            return (double)bigInt;
        }

        // Handle decimal with high precision
        if (value is decimal dec)
            return (double)dec;

        // Handle other numeric types that might cause issues
        var type = value.GetType();
        if (type.FullName?.Contains("HugeInt") == true ||
            type.FullName?.Contains("Int128") == true)
        {
            // Try to convert via ToString and parse
            if (double.TryParse(value.ToString(), out var d))
                return d;
        }

        return value;
    }

    /// <summary>
    /// Append a value to a DuckDB appender row, handling type dispatch.
    /// </summary>
    private static void AppendTypedValue(DuckDB.NET.Data.IDuckDBAppenderRow row, object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            row.AppendNullValue();
            return;
        }

        switch (value)
        {
            case bool b: row.AppendValue(b); break;
            case byte b: row.AppendValue(b); break;
            case sbyte sb: row.AppendValue(sb); break;
            case short s: row.AppendValue(s); break;
            case ushort us: row.AppendValue(us); break;
            case int i: row.AppendValue(i); break;
            case uint ui: row.AppendValue(ui); break;
            case long l: row.AppendValue(l); break;
            case ulong ul: row.AppendValue(ul); break;
            case float f: row.AppendValue(f); break;
            case double d: row.AppendValue(d); break;
            case decimal dec: row.AppendValue(dec); break;
            case string str: row.AppendValue(str); break;
            case DateTime dt: row.AppendValue(dt); break;
            case DateOnly date: row.AppendValue(date); break;
            case TimeOnly time: row.AppendValue(time); break;
            case byte[] bytes: row.AppendValue(bytes); break;
            case System.Numerics.BigInteger bigInt:
                // Convert BigInteger to long or double
                if (bigInt >= long.MinValue && bigInt <= long.MaxValue)
                    row.AppendValue((long)bigInt);
                else
                    row.AppendValue((double)bigInt);
                break;
            default:
                // Fallback: convert to string
                row.AppendValue(value.ToString() ?? "");
                break;
        }
    }

    /// <summary>
    /// Map .NET types to DuckDB column types.
    /// </summary>
    private static string MapTypeToDuckDB(Type type)
    {
        if (type == typeof(int) || type == typeof(int?)) return "INTEGER";
        if (type == typeof(long) || type == typeof(long?)) return "BIGINT";
        if (type == typeof(short) || type == typeof(short?)) return "SMALLINT";
        if (type == typeof(byte) || type == typeof(byte?)) return "TINYINT";
        if (type == typeof(float) || type == typeof(float?)) return "FLOAT";
        if (type == typeof(double) || type == typeof(double?)) return "DOUBLE";
        if (type == typeof(decimal) || type == typeof(decimal?)) return "DECIMAL";
        if (type == typeof(bool) || type == typeof(bool?)) return "BOOLEAN";
        if (type == typeof(string)) return "VARCHAR";
        if (type == typeof(DateTime) || type == typeof(DateTime?)) return "TIMESTAMP";
        if (type == typeof(DateOnly) || type == typeof(DateOnly?)) return "DATE";
        if (type == typeof(TimeOnly) || type == typeof(TimeOnly?)) return "TIME";
        if (type == typeof(byte[])) return "BLOB";
        return "VARCHAR";
    }

    /// <summary>
    /// Drop temp tables created during query resolution.
    /// </summary>
    private static void CleanupTempTables(List<string> tempTables)
    {
        var conn = GetConnection();
        foreach (var tableName in tempTables)
        {
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"DROP TABLE IF EXISTS \"{tableName}\"";
                cmd.ExecuteNonQuery();
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
}
