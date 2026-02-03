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
    private const int DuckOutMaxRows = 200_000;

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

    internal static DuckDBConnection GetConnection()
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

    /// <summary>
    /// Drop a DuckDB temp table. Called when a handle is evicted from ResultStore.
    /// </summary>
    internal static void DropTempTable(string tableName)
    {
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"DROP TABLE IF EXISTS \"{tableName}\"";
            cmd.ExecuteNonQuery();
            System.Diagnostics.Debug.WriteLine($"[XlDuck] Dropped temp table: {tableName}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[XlDuck] Error dropping table {tableName}: {ex.Message}");
        }
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
                return QueryTableToArray(stored);
            }
            else if (FragmentStore.IsHandle(handle))
            {
                // Execute the fragment and output results directly (no temp table needed)
                var fragment = FragmentStore.Get(handle);
                if (fragment == null)
                {
                    return new object[,] { { FormatError("notfound", $"Fragment not found: {handle}") } };
                }

                var (resolvedSql, referencedHandles) = ResolveParameters(fragment.Sql, fragment.Args, new HashSet<string> { handle });
                try
                {
                    return ExecuteAndReturnArray(resolvedSql);
                }
                finally
                {
                    DecrementHandleRefCounts(referencedHandles);
                }
            }
            else
            {
                return new object[,] { { FormatError("invalid", $"Invalid handle format: {handle}") } };
            }
        }
        catch (Exception ex)
        {
            return new object[,] { { FormatException(ex) } };
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
            var (resolvedSql, referencedHandles) = ResolveParameters(sql, args, new HashSet<string>());
            try
            {
                return ExecuteAndReturnArray(resolvedSql);
            }
            finally
            {
                DecrementHandleRefCounts(referencedHandles);
            }
        }
        catch (Exception ex)
        {
            return new object[,] { { FormatException(ex) } };
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
    /// Execute a query, store the result as a DuckDB temp table, and return the handle. Called by RTD server.
    /// </summary>
    internal static string ExecuteQueryInternal(string sql, object[] args)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var (resolvedSql, referencedHandles) = ResolveParameters(sql, args, new HashSet<string>());
        var resolveTime = sw.ElapsedMilliseconds;

        try
        {
            var conn = GetConnection();
            var duckTableName = $"_xlduck_res_{Guid.NewGuid():N}";

            // Create temp table with query results
            sw.Restart();
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"CREATE TEMP TABLE \"{duckTableName}\" AS {resolvedSql}";
                cmd.ExecuteNonQuery();
            }
            var createTime = sw.ElapsedMilliseconds;

            // Get schema from PRAGMA table_info
            var columnNames = GetTableColumnNames(conn, duckTableName);

            // Get row count
            sw.Restart();
            long rowCount;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"SELECT COUNT(*) FROM \"{duckTableName}\"";
                rowCount = Convert.ToInt64(cmd.ExecuteScalar());
            }
            var countTime = sw.ElapsedMilliseconds;

            var stored = new StoredResult(duckTableName, columnNames, rowCount);
            var handle = ResultStore.Store(stored);

            System.Diagnostics.Debug.WriteLine($"[DuckQuery] resolve={resolveTime}ms create={createTime}ms count={countTime}ms rows={rowCount} cols={columnNames.Length}");
            return handle;
        }
        finally
        {
            DecrementHandleRefCounts(referencedHandles);
        }
    }

    /// <summary>
    /// Create a fragment, validate it, and return the handle. Called by RTD server.
    /// </summary>
    internal static string CreateFragmentInternal(string sql, object[] args)
    {
        // Validate the SQL by resolving parameters and running EXPLAIN
        var (resolvedSql, referencedHandles) = ResolveParameters(sql, args, new HashSet<string>());
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"EXPLAIN {resolvedSql}";
            cmd.ExecuteNonQuery();
        }
        finally
        {
            DecrementHandleRefCounts(referencedHandles);
        }

        // Store the fragment with original SQL and args
        var fragment = new StoredFragment(sql, args);
        return FragmentStore.Store(fragment);
    }

    /// <summary>
    /// Get column names from a table using PRAGMA table_info.
    /// </summary>
    private static string[] GetTableColumnNames(DuckDBConnection conn, string tableName)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info('{tableName}')";
        using var reader = cmd.ExecuteReader();

        var names = new List<string>();
        while (reader.Read())
        {
            names.Add(reader.GetString(reader.GetOrdinal("name")));
        }
        return names.ToArray();
    }

    /// <summary>
    /// Query a stored result table and return as Excel array with limit and truncation footer.
    /// </summary>
    private static object[,] QueryTableToArray(StoredResult stored)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var conn = GetConnection();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM \"{stored.DuckTableName}\" LIMIT {DuckOutMaxRows + 1}";
        using var reader = cmd.ExecuteReader();

        var cols = stored.ColumnNames.Length;
        if (cols == 0)
        {
            return new object[,] { { FormatError("query", "No columns") } };
        }

        var rows = new List<object?[]>(Math.Min((int)stored.RowCount + 1, DuckOutMaxRows + 1));
        while (reader.Read())
        {
            var row = new object?[cols];
            for (int j = 0; j < cols; j++)
            {
                row[j] = reader.IsDBNull(j) ? null : reader.GetValue(j);
            }
            rows.Add(row);
        }
        var readTime = sw.ElapsedMilliseconds;

        var truncated = rows.Count > DuckOutMaxRows;
        var dataRowsToEmit = truncated ? DuckOutMaxRows : rows.Count;

        // +1 for header, +1 for footer if truncated
        var outRows = 1 + dataRowsToEmit + (truncated ? 1 : 0);
        var result = new object[outRows, cols];

        // Header row
        for (int j = 0; j < cols; j++)
        {
            result[0, j] = stored.ColumnNames[j];
        }

        // Data rows
        for (int i = 0; i < dataRowsToEmit; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                result[i + 1, j] = ConvertToExcelValue(rows[i]![j]);
            }
        }

        // Footer if truncated
        if (truncated)
        {
            result[1 + dataRowsToEmit, 0] = $"(Truncated) Showing first {DuckOutMaxRows:N0} of {stored.RowCount:N0} rows";
            for (int j = 1; j < cols; j++)
            {
                result[1 + dataRowsToEmit, j] = "";
            }
        }

        System.Diagnostics.Debug.WriteLine($"[DuckOut] read={readTime}ms rows={dataRowsToEmit} cols={cols} truncated={truncated}");
        return result;
    }

    /// <summary>
    /// Execute a query and return results directly as an Excel array (for DuckQueryOut and fragment execution).
    /// </summary>
    private static object[,] ExecuteAndReturnArray(string sql)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        var conn = GetConnection();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        using var reader = cmd.ExecuteReader();

        var fieldCount = reader.FieldCount;
        if (fieldCount == 0)
        {
            return new object[,] { { FormatError("query", "No columns") } };
        }

        var columnNames = new string[fieldCount];
        for (int i = 0; i < fieldCount; i++)
        {
            columnNames[i] = reader.GetName(i);
        }

        var rows = new List<object?[]>(DuckOutMaxRows + 1);
        while (reader.Read() && rows.Count <= DuckOutMaxRows)
        {
            var row = new object?[fieldCount];
            for (int j = 0; j < fieldCount; j++)
            {
                row[j] = reader.IsDBNull(j) ? null : reader.GetValue(j);
            }
            rows.Add(row);
        }
        var readTime = sw.ElapsedMilliseconds;

        // Check if there are more rows (we read maxRows + 1 to detect truncation)
        var truncated = rows.Count > DuckOutMaxRows;
        var dataRowsToEmit = truncated ? DuckOutMaxRows : rows.Count;

        var outRows = 1 + dataRowsToEmit + (truncated ? 1 : 0);
        var result = new object[outRows, fieldCount];

        // Header row
        for (int j = 0; j < fieldCount; j++)
        {
            result[0, j] = columnNames[j];
        }

        // Data rows
        for (int i = 0; i < dataRowsToEmit; i++)
        {
            for (int j = 0; j < fieldCount; j++)
            {
                result[i + 1, j] = ConvertToExcelValue(rows[i]![j]);
            }
        }

        // Footer if truncated
        if (truncated)
        {
            result[1 + dataRowsToEmit, 0] = $"(Truncated) Showing first {DuckOutMaxRows:N0} rows";
            for (int j = 1; j < fieldCount; j++)
            {
                result[1 + dataRowsToEmit, j] = "";
            }
        }

        System.Diagnostics.Debug.WriteLine($"[DuckQueryOut] read={readTime}ms rows={dataRowsToEmit} cols={fieldCount} truncated={truncated}");
        return result;
    }

    /// <summary>
    /// Collect optional name/value pairs into an array, filtering out missing values.
    /// The @config sentinel is treated specially - added standalone without a value.
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

            // @config sentinel is standalone - don't add its (empty) value
            if (name is string nameStr && nameStr == ConfigSentinel)
            {
                result.Add(name);
                continue;
            }

            result.Add(name);
            result.Add(value);
        }
        return result.ToArray();
    }

    /// <summary>
    /// Parse SQL for :name placeholders, look up handles, and return resolved SQL.
    /// Table handles are resolved to their DuckDB table names.
    /// Fragment handles are resolved recursively and inlined as subqueries.
    /// Returns list of table handles that were referenced (their refcounts were incremented).
    /// </summary>
    private static (string resolvedSql, List<string> referencedHandles) ResolveParameters(string sql, object[] args, HashSet<string> visitedFragments)
    {
        var referencedHandles = new List<string>();

        if (args.Length == 0)
        {
            return (sql, referencedHandles);
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
                // Table handle: reference the existing DuckDB temp table directly
                var stored = ResultStore.Get(value) ?? throw new ArgumentException($"Handle not found: {value}");

                // Increment refcount to prevent table from being dropped during query
                ResultStore.IncrementRefCount(value);
                referencedHandles.Add(value);

                parameters[name] = $"\"{stored.DuckTableName}\"";
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
                var (resolvedFragmentSql, fragmentReferencedHandles) = ResolveParameters(fragment.Sql, fragment.Args, visitedFragments);
                visitedFragments.Remove(value);

                // Collect any referenced handles from fragment resolution
                referencedHandles.AddRange(fragmentReferencedHandles);

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

        return (resolvedSql, referencedHandles);
    }

    /// <summary>
    /// Decrement refcounts for handles that were referenced during a query.
    /// </summary>
    private static void DecrementHandleRefCounts(List<string> handles)
    {
        foreach (var handle in handles)
        {
            var evicted = ResultStore.DecrementRefCount(handle);
            if (evicted != null)
            {
                DropTempTable(evicted.DuckTableName);
            }
        }
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
}
