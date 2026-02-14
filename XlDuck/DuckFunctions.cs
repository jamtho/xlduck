// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Collections.Concurrent;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
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
        Log.Write("[AddIn] AutoOpen called - add-in loading");
        System.Diagnostics.Debug.WriteLine("[XlDuck] Add-in loaded");
    }

    public void AutoClose()
    {
        Log.Write("[AddIn] AutoClose called - add-in unloading");
        System.Diagnostics.Debug.WriteLine("[XlDuck] Add-in unloaded");
    }
}

public static class DuckFunctions
{
    private const string Version = "0.1";
    private const int DuckOutMaxRows = 200_000;

    private static DuckDBConnection? _connection;
    private static readonly object _connLock = new();
    private static volatile int _interruptEpoch;
    private static readonly object _queryLock = new();
    [ThreadStatic] private static int _threadEpoch;

    private static volatile bool _queriesPaused;
    private static readonly ManualResetEventSlim _unpauseEvent = new(true);

    /// <summary>
    /// Monotonically increasing epoch, bumped on each Interrupt() call.
    /// Query threads capture this before executing and bail if it changes.
    /// </summary>
    internal static int InterruptEpoch => _interruptEpoch;

    /// <summary>
    /// Set the interrupt epoch for the current thread. Called by RTD threads
    /// before executing so ThrowIfInterrupted can detect stale queries.
    /// </summary>
    internal static void SetThreadEpoch(int epoch) => _threadEpoch = epoch;

    /// <summary>
    /// Throw if an interrupt has occurred since the current thread's epoch was set.
    /// </summary>
    private static void ThrowIfInterrupted()
    {
        if (_threadEpoch != _interruptEpoch)
            throw new OperationCanceledException("Query cancelled");
    }

    // Ready flag - DuckQueryAfterConfig/DuckFragAfterConfig wait until DuckConfigReady() is called
    internal static bool IsReady { get; private set; } = false;

    // Stash for pending DuckCapture data (hash → array), consumed by RTD ConnectData
    private static readonly ConcurrentDictionary<string, object[,]> _pendingCaptures = new();

    // Full error messages keyed by error ID (RTD truncates values to 255 chars)
    private static readonly ConcurrentDictionary<long, (string Category, string Message)> _fullErrors = new();
    private static long _nextErrorId;

    // Status URL prefixes (# prefix follows Excel convention)
    internal const string BlockedPrefix = "#duck://blocked/";
    internal const string ErrorPrefix = "#duck://error/";
    internal const string ConfigBlockedStatus = "#duck://blocked/config|Waiting for DuckConfigReady()";
    internal const string PausedBlockedStatus = "#duck://blocked/paused|Queries paused";

    /// <summary>
    /// Format an error message as a duck:// URL.
    /// Error ID is embedded so the preview pane can look up the full message
    /// (RTD truncates return values to 255 characters).
    /// Format: #duck://error/ID/category|message
    /// </summary>
    internal static string FormatError(string category, string message)
    {
        // Remove newlines for single-line cell display
        var cleanMessage = message.Replace("\r", "").Replace("\n", " ");
        var id = Interlocked.Increment(ref _nextErrorId);
        _fullErrors[id] = (category, cleanMessage);
        return $"{ErrorPrefix}{id}/{category}|{cleanMessage}";
    }

    /// <summary>
    /// Look up full error details by ID. Used by preview pane to bypass RTD 255-char truncation.
    /// </summary>
    internal static (string Category, string Message)? GetFullError(long id)
    {
        return _fullErrors.TryGetValue(id, out var error) ? error : null;
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

    internal static void Interrupt()
    {
        Interlocked.Increment(ref _interruptEpoch);
        var conn = _connection;
        if (conn == null) return;
        conn.NativeConnection.Interrupt();
        Log.Write("[DuckFunctions] Query interrupted by user");
    }

    [ExcelCommand(Name = "DuckInterrupt")]
    public static void DuckInterruptCommand()
    {
        Interrupt();
    }

    internal static bool QueriesPaused => _queriesPaused;

    internal static void SetQueriesPaused(bool paused)
    {
        if (paused)
        {
            _queriesPaused = true;
            _unpauseEvent.Reset();
            Interrupt();
            Log.Write("[DuckFunctions] Queries paused");
        }
        else
        {
            _queriesPaused = false;
            _unpauseEvent.Set();
            Log.Write("[DuckFunctions] Queries resumed");
        }
    }

    internal static bool WaitForUnpause(CancellationToken ct = default)
    {
        while (_queriesPaused)
        {
            try { _unpauseEvent.Wait(ct); }
            catch (OperationCanceledException) { return false; }
        }
        return true;
    }

    [ExcelCommand(Name = "DuckPauseQueries")]
    public static void DuckPauseQueriesCommand() => SetQueriesPaused(true);

    [ExcelCommand(Name = "DuckResumeQueries")]
    public static void DuckResumeQueriesCommand() => SetQueriesPaused(false);

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
        Log.Write("[DuckVersion] Called");
        return Version;
    }

    [ExcelFunction(Description = "Signal that configuration is complete. DuckQueryAfterConfig/DuckFragAfterConfig wait until this is called.")]
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

    [ExcelFunction(Description = "Convert an Excel date serial number to a SQL date string (yyyy-MM-dd).")]
    public static string DuckDate(
        [ExcelArgument(Description = "Cell containing a date")] double value)
    {
        return DateTime.FromOADate(value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    [ExcelFunction(Description = "Convert an Excel date/time serial number to a SQL datetime string (yyyy-MM-dd HH:mm:ss).")]
    public static string DuckDateTime(
        [ExcelArgument(Description = "Cell containing a date/time")] double value)
    {
        return DateTime.FromOADate(value).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL query and return a handle. Use ? placeholders for positional arguments.")]
    public static object DuckQuery(
        [ExcelArgument(Description = "SQL query with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
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

    [ExcelFunction(Description = "Execute a DuckDB SQL query after DuckConfigReady() is called. Use ? placeholders for positional arguments.")]
    public static object DuckQueryAfterConfig(
        [ExcelArgument(Description = "SQL query with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
            // Build topic info: ["query-config", sql, arg1, arg2, ...]
            var topicInfo = new List<string> { "query-config", sql };
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

    [ExcelFunction(Description = "Execute a DuckDB SQL query and output results as a spilled array. Use ? placeholders for positional arguments.")]
    public static object[,] DuckQueryOut(
        [ExcelArgument(Description = "SQL query with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
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

    [ExcelFunction(Description = "Execute a DuckDB SQL query and return a single value (first column, first row). Use ? placeholders for positional arguments.")]
    public static object DuckQueryOutScalar(
        [ExcelArgument(Description = "SQL query with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
            var (resolvedSql, referencedHandles) = ResolveParameters(sql, args, new HashSet<string>());
            try
            {
                var conn = GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = resolvedSql;
                var result = cmd.ExecuteScalar();
                return ConvertToExcelValue(result);
            }
            finally
            {
                DecrementHandleRefCounts(referencedHandles);
            }
        }
        catch (Exception ex)
        {
            return FormatException(ex);
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

    [ExcelFunction(Description = "Create a chart from data. Templates: bar, line, point, area, histogram, heatmap, boxplot. Overrides: x, y, color, label, tooltip, title, value, xmin, xmax, ymin, ymax.")]
    public static object DuckPlot(
        [ExcelArgument(Description = "Data handle (table or fragment)")] string dataHandle,
        [ExcelArgument(Description = "bar, line, point, area, histogram, heatmap, boxplot")] string template,
        [ExcelArgument(Description = "bar/line/point/area: x,y. histogram: x. heatmap: x,y,value. boxplot: x,y. All: color, label, tooltip, title, xmin, xmax, ymin, ymax")] object name1 = null!,
        [ExcelArgument(Description = "Column name or literal for this override")] object value1 = null!,
        [ExcelArgument(Description = "Override name (see name1 for allowed names)")] object name2 = null!,
        [ExcelArgument(Description = "Column name or literal for this override")] object value2 = null!,
        [ExcelArgument(Description = "Override name")] object name3 = null!,
        [ExcelArgument(Description = "Override value")] object value3 = null!,
        [ExcelArgument(Description = "Override name")] object name4 = null!,
        [ExcelArgument(Description = "Override value")] object value4 = null!,
        [ExcelArgument(Description = "Override name")] object name5 = null!,
        [ExcelArgument(Description = "Override value")] object value5 = null!,
        [ExcelArgument(Description = "Override name")] object name6 = null!,
        [ExcelArgument(Description = "Override value")] object value6 = null!,
        [ExcelArgument(Description = "Override name")] object name7 = null!,
        [ExcelArgument(Description = "Override value")] object value7 = null!,
        [ExcelArgument(Description = "Override name")] object name8 = null!,
        [ExcelArgument(Description = "Override value")] object value8 = null!)
    {
        try
        {
            // Validate template
            if (!PlotTemplates.IsValidTemplate(template))
            {
                return FormatError("invalid", $"Unknown template: {template}. Valid: {string.Join(", ", PlotTemplates.TemplateNames)}");
            }

            // Validate data handle
            if (!ResultStore.IsHandle(dataHandle) && !FragmentStore.IsHandle(dataHandle))
            {
                // Check for error/blocked status
                if (IsErrorOrBlocked(dataHandle))
                {
                    return dataHandle; // Pass through error/blocked status
                }
                return FormatError("invalid", "Data must be a table or fragment handle");
            }

            var args = CollectArgs(name1, value1, name2, value2, name3, value3, name4, value4,
                                   name5, value5, name6, value6, name7, value7, name8, value8);

            // Validate required overrides
            var overrides = new Dictionary<string, string>();
            for (int i = 0; i + 1 < args.Length; i += 2)
            {
                var name = args[i]?.ToString() ?? "";
                var value = args[i + 1]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(name))
                {
                    overrides[name] = value;
                }
            }

            if (!overrides.ContainsKey("x"))
                return FormatError("invalid", "Missing required override: x");

            // y is required for most templates, but not histogram
            if (template != "histogram" && !overrides.ContainsKey("y"))
                return FormatError("invalid", "Missing required override: y");

            // heatmap requires value for color intensity
            if (template == "heatmap" && !overrides.ContainsKey("value"))
                return FormatError("invalid", "Missing required override: value (for color intensity)");

            // Build topic info: ["plot", dataHandle, template, arg1, arg2, ...]
            var topicInfo = new List<string> { "plot", dataHandle, template };
            topicInfo.AddRange(args.Select(a => a?.ToString() ?? ""));

            return XlCall.RTD("XlDuck.DuckRtdServer", null, topicInfo.ToArray());
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    [ExcelFunction(Description = "Create a SQL fragment for lazy evaluation. Use ? placeholders for positional arguments.")]
    public static object DuckFrag(
        [ExcelArgument(Description = "SQL query (SELECT or PIVOT) with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
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

    [ExcelFunction(Description = "Create a SQL fragment after DuckConfigReady() is called. Use ? placeholders for positional arguments.")]
    public static object DuckFragAfterConfig(
        [ExcelArgument(Description = "SQL query (SELECT or PIVOT) with optional ? placeholders")] string sql,
        [ExcelArgument(Description = "First argument (replaces first ?)")] object arg1 = null!,
        [ExcelArgument(Description = "Second argument (replaces second ?)")] object arg2 = null!,
        [ExcelArgument(Description = "Third argument")] object arg3 = null!,
        [ExcelArgument(Description = "Fourth argument")] object arg4 = null!,
        [ExcelArgument(Description = "Fifth argument")] object arg5 = null!,
        [ExcelArgument(Description = "Sixth argument")] object arg6 = null!,
        [ExcelArgument(Description = "Seventh argument")] object arg7 = null!,
        [ExcelArgument(Description = "Eighth argument")] object arg8 = null!)
    {
        try
        {
            var args = CollectArgs(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
            // Build topic info: ["frag-config", sql, arg1, arg2, ...]
            var topicInfo = new List<string> { "frag-config", sql };
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
            lock (_queryLock)
            {
                ThrowIfInterrupted();

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

                var stored = new StoredResult(duckTableName, columnNames, rowCount, sql, args);
                var handle = ResultStore.Store(stored);

                System.Diagnostics.Debug.WriteLine($"[DuckQuery] resolve={resolveTime}ms create={createTime}ms count={countTime}ms rows={rowCount} cols={columnNames.Length}");
                return handle;
            }
        }
        finally
        {
            DecrementHandleRefCounts(referencedHandles);
        }
    }

    /// <summary>
    /// Create a plot configuration and return the handle. Called by RTD server.
    /// </summary>
    internal static string CreatePlotInternal(string dataHandle, string template, object[] args)
    {
        // Build overrides dictionary
        var overrides = new Dictionary<string, string>();
        for (int i = 0; i + 1 < args.Length; i += 2)
        {
            var name = args[i]?.ToString() ?? "";
            var value = args[i + 1]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(name))
            {
                overrides[name] = value;
            }
        }

        var plot = new StoredPlot(dataHandle, template, overrides);
        return PlotStore.Store(plot);
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
            lock (_queryLock)
            {
                ThrowIfInterrupted();

                var conn = GetConnection();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $"EXPLAIN {resolvedSql}";
                cmd.ExecuteNonQuery();
            }
        }
        finally
        {
            DecrementHandleRefCounts(referencedHandles);
        }

        // Store the fragment with original SQL and args
        var fragment = new StoredFragment(sql, args);
        return FragmentStore.Store(fragment);
    }

    // ─── DuckCapture ───────────────────────────────────────────────

    [ExcelFunction(Description = "Capture a sheet range as a DuckDB table. First row = headers.")]
    public static object DuckCapture(
        [ExcelArgument(Description = "Range (first row = headers, rest = data)")] object[,] range)
    {
        try
        {
            var rows = range.GetLength(0);
            var cols = range.GetLength(1);

            if (rows < 2 || cols < 1)
                return FormatError("invalid", "Range must have at least 1 header row and 1 data row");

            var hash = ComputeRangeHash(range);
            _pendingCaptures[hash] = range;

            return XlCall.RTD("XlDuck.DuckRtdServer", null, "capture", hash);
        }
        catch (Exception ex)
        {
            return FormatException(ex);
        }
    }

    /// <summary>
    /// Remove and return stashed capture data by hash. Called by RTD server.
    /// </summary>
    internal static object[,]? TakePendingCapture(string hash)
    {
        _pendingCaptures.TryRemove(hash, out var data);
        return data;
    }

    /// <summary>
    /// Capture a range array into a DuckDB temp table and return the handle. Called by RTD server.
    /// </summary>
    internal static string CaptureRangeInternal(object[,] data)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);
        var dataRows = rows - 1;

        var headers = ExtractHeaders(data, cols);
        var types = InferColumnTypes(data, dataRows, cols);

        return CreateCaptureTable(headers, types, data, dataRows, cols);
    }

    /// <summary>
    /// Compute SHA256 hash of range dimensions and all cell values.
    /// </summary>
    private static string ComputeRangeHash(object[,] data)
    {
        var rows = data.GetLength(0);
        var cols = data.GetLength(1);

        using var sha = SHA256.Create();
        var sb = new StringBuilder();
        sb.Append(rows).Append('x').Append(cols).Append('\0');

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                sb.Append(ConvertCellToString(data[r, c])).Append('\0');
            }
        }

        var hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
        return Convert.ToHexString(hashBytes);
    }

    /// <summary>
    /// Extract and sanitize column headers from row 0.
    /// Deduplicates by appending _2, _3, etc.
    /// </summary>
    private static string[] ExtractHeaders(object[,] data, int cols)
    {
        var headers = new string[cols];
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int c = 0; c < cols; c++)
        {
            var raw = ConvertCellToString(data[0, c]).Trim();
            if (string.IsNullOrEmpty(raw))
                raw = $"col{c + 1}";

            // Sanitize: remove characters that are problematic in SQL identifiers
            raw = Regex.Replace(raw, @"[^\w]", "_");
            if (raw.Length == 0 || char.IsDigit(raw[0]))
                raw = "_" + raw;

            if (seen.TryGetValue(raw, out var count))
            {
                seen[raw] = count + 1;
                raw = $"{raw}_{count + 1}";
            }
            else
            {
                seen[raw] = 1;
            }

            headers[c] = raw;
        }

        return headers;
    }

    private enum CaptureType { Double, Boolean, Varchar }

    /// <summary>
    /// Infer column types by scanning data rows.
    /// </summary>
    private static CaptureType[] InferColumnTypes(object[,] data, int dataRows, int cols)
    {
        var types = new CaptureType[cols];

        for (int c = 0; c < cols; c++)
        {
            bool allDouble = true;
            bool allBool = true;
            bool hasValue = false;

            for (int r = 1; r <= dataRows; r++)
            {
                var cell = data[r, c];
                if (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
                    continue;

                hasValue = true;

                if (cell is not double)
                    allDouble = false;
                if (cell is not bool)
                    allBool = false;

                if (!allDouble && !allBool)
                    break;
            }

            if (!hasValue || (!allDouble && !allBool))
                types[c] = CaptureType.Varchar;
            else if (allDouble)
                types[c] = CaptureType.Double;
            else
                types[c] = CaptureType.Boolean;
        }

        return types;
    }

    private static string DuckDbTypeName(CaptureType type) => type switch
    {
        CaptureType.Double => "DOUBLE",
        CaptureType.Boolean => "BOOLEAN",
        _ => "VARCHAR"
    };

    /// <summary>
    /// Create the DuckDB temp table, insert data, store in ResultStore, return handle.
    /// </summary>
    private static string CreateCaptureTable(string[] headers, CaptureType[] types, object[,] data, int dataRows, int cols)
    {
        lock (_queryLock)
        {
            ThrowIfInterrupted();

            var conn = GetConnection();
            var tableName = $"_xlduck_cap_{Guid.NewGuid():N}";

            // CREATE TEMP TABLE
            var colDefs = new StringBuilder();
            for (int c = 0; c < cols; c++)
            {
                if (c > 0) colDefs.Append(", ");
                colDefs.Append('"').Append(headers[c]).Append("\" ").Append(DuckDbTypeName(types[c]));
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = $"CREATE TEMP TABLE \"{tableName}\" ({colDefs})";
                cmd.ExecuteNonQuery();
            }

            // INSERT in batches of 1000
            const int batchSize = 1000;
            for (int batchStart = 0; batchStart < dataRows; batchStart += batchSize)
            {
                var batchEnd = Math.Min(batchStart + batchSize, dataRows);
                var sb = new StringBuilder();
                sb.Append($"INSERT INTO \"{tableName}\" VALUES ");

                for (int r = batchStart; r < batchEnd; r++)
                {
                    if (r > batchStart) sb.Append(", ");
                    sb.Append('(');
                    for (int c = 0; c < cols; c++)
                    {
                        if (c > 0) sb.Append(", ");
                        sb.Append(FormatCellAsSqlLiteral(data[r + 1, c], types[c]));
                    }
                    sb.Append(')');
                }

                using var cmd = conn.CreateCommand();
                cmd.CommandText = sb.ToString();
                cmd.ExecuteNonQuery();
            }

            var stored = new StoredResult(tableName, headers, dataRows, "(captured from Excel range)");
            return ResultStore.Store(stored);
        }
    }

    /// <summary>
    /// Format a cell value as a SQL literal based on inferred column type.
    /// </summary>
    private static string FormatCellAsSqlLiteral(object cell, CaptureType colType)
    {
        if (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            return "NULL";

        return colType switch
        {
            CaptureType.Double when cell is double d => d.ToString(CultureInfo.InvariantCulture),
            CaptureType.Boolean when cell is bool b => b ? "TRUE" : "FALSE",
            _ => $"'{ConvertCellToString(cell).Replace("'", "''")}'"
        };
    }

    /// <summary>
    /// Convert a cell value to its string representation.
    /// </summary>
    private static string ConvertCellToString(object cell)
    {
        if (cell == null || cell is ExcelEmpty || cell is ExcelMissing)
            return "";
        if (cell is ExcelError err)
            return $"#ERR:{err}";
        if (cell is double d)
            return d.ToString(CultureInfo.InvariantCulture);
        if (cell is bool b)
            return b ? "TRUE" : "FALSE";
        return cell.ToString() ?? "";
    }

    // ─── End DuckCapture ────────────────────────────────────────────

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
    /// Collect optional positional values into an array, stopping at the first missing value.
    /// </summary>
    private static object[] CollectArgs(params object[] values)
    {
        var result = new List<object>();
        foreach (var value in values)
        {
            if (value == null || value is ExcelMissing || value is ExcelEmpty)
                break;
            if (value is string s && string.IsNullOrEmpty(s))
                break;
            result.Add(value);
        }
        return result.ToArray();
    }

    /// <summary>
    /// Replace ? placeholders with positional argument values.
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

        // Resolve each positional value
        var resolvedValues = new List<string>();
        foreach (var arg in args)
        {
            var value = arg?.ToString() ?? "";

            if (ResultStore.IsHandle(value))
            {
                // Table handle: reference the existing DuckDB temp table directly
                var stored = ResultStore.Get(value) ?? throw new ArgumentException($"Handle not found: {value}");

                // Increment refcount to prevent table from being dropped during query
                ResultStore.IncrementRefCount(value);
                referencedHandles.Add(value);

                resolvedValues.Add($"\"{stored.DuckTableName}\"");
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
                resolvedValues.Add($"({resolvedFragmentSql})");
            }
            else
            {
                // Quote string values for SQL (escape single quotes)
                var escaped = value.Replace("'", "''");
                resolvedValues.Add($"'{escaped}'");
            }
        }

        // Replace ? placeholders left-to-right with resolved values
        int paramIndex = 0;
        var resolvedSql = Regex.Replace(sql, @"\?", match =>
        {
            if (paramIndex >= resolvedValues.Count)
                throw new ArgumentException($"More ? placeholders than arguments ({resolvedValues.Count} provided)");
            return resolvedValues[paramIndex++];
        });
        if (paramIndex < resolvedValues.Count)
            throw new ArgumentException($"More arguments ({resolvedValues.Count}) than ? placeholders ({paramIndex})");

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

        // Handle date/time types that COM interop cannot marshal
        if (value is DateOnly dateOnly)
            return new DateTime(dateOnly.Year, dateOnly.Month, dateOnly.Day).ToOADate();
        if (value is TimeOnly timeOnly)
            return timeOnly.ToTimeSpan().TotalDays;
        if (value is DateTime dt)
            return dt.ToOADate();
        if (value is DateTimeOffset dto)
            return dto.DateTime.ToOADate();

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
