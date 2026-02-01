using System.Text.RegularExpressions;
using ExcelDna.Integration;
using DuckDB.NET.Data;

namespace XlDuck;

public static class DuckFunctions
{
    private static DuckDBConnection? _connection;
    private static readonly object _connLock = new();

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
            var (resolvedSql, tempTables) = ResolveParameters(sql, args);
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

                var storedResult = new StoredResult(columnNames, columnTypes, rows);
                var handle = ResultStore.Store(storedResult);
                return handle;
            }
            finally
            {
                CleanupTempTables(tempTables);
            }
        }
        catch (Exception ex)
        {
            return $"#ERROR: {ex.Message}";
        }
    }

    [ExcelFunction(Description = "Output a query result as a spilled array with headers.")]
    public static object[,] DuckQueryOut(
        [ExcelArgument(Description = "Handle from DuckQuery (e.g. duck://t/1)")] string handle)
    {
        try
        {
            var stored = ResultStore.Get(handle);
            if (stored == null)
            {
                return new object[,] { { $"#ERROR: Handle not found: {handle}" } };
            }

            var fieldCount = stored.ColumnNames.Length;
            var rowCount = stored.Rows.Count;

            if (fieldCount == 0)
            {
                return new object[,] { { "#ERROR: No columns" } };
            }

            var result = new object[rowCount + 1, fieldCount];

            // Header row
            for (int j = 0; j < fieldCount; j++)
            {
                result[0, j] = stored.ColumnNames[j];
            }

            // Data rows
            for (int i = 0; i < rowCount; i++)
            {
                var row = stored.Rows[i];
                for (int j = 0; j < fieldCount; j++)
                {
                    result[i + 1, j] = ConvertToExcelValue(row[j]);
                }
            }

            return result;
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
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            var rowsAffected = cmd.ExecuteNonQuery();
            return $"OK ({rowsAffected} rows affected)";
        }
        catch (Exception ex)
        {
            return $"#ERROR: {ex.Message}";
        }
    }

    [ExcelFunction(Description = "Get the DuckDB version")]
    public static string DuckVersion()
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
            return $"#ERROR: {ex.Message}";
        }
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
    /// Parse SQL for :name placeholders, look up handles, create temp tables, and return resolved SQL.
    /// </summary>
    private static (string resolvedSql, List<string> tempTables) ResolveParameters(string sql, object[] args)
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
                var stored = ResultStore.Get(value) ?? throw new ArgumentException($"Handle not found: {value}");
                var tempTableName = CreateTempTable(stored);
                tempTables.Add(tempTableName);
                parameters[name] = tempTableName;
            }
            else
            {
                parameters[name] = value;
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
            var placeholders = string.Join(", ", Enumerable.Range(0, stored.ColumnNames.Length).Select(i => $"${i + 1}"));
            var insertSql = $"INSERT INTO \"{tableName}\" VALUES ({placeholders})";

            foreach (var row in stored.Rows)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = insertSql;
                for (int i = 0; i < row.Length; i++)
                {
                    var param = cmd.CreateParameter();
                    param.Value = row[i] ?? DBNull.Value;
                    cmd.Parameters.Add(param);
                }
                cmd.ExecuteNonQuery();
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
