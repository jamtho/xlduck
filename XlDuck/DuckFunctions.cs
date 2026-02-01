using ExcelDna.Integration;
using DuckDB.NET.Data;

namespace XlDuck;

public static class DuckFunctions
{
    // Simple test function - no DuckDB dependency
    [ExcelFunction(Description = "Test function - returns Hello")]
    public static string XlDuckHello()
    {
        return "Hello from XlDuck!";
    }

    private static DuckDBConnection? _connection;
    private static readonly object _lock = new();

    private static DuckDBConnection GetConnection()
    {
        if (_connection == null)
        {
            lock (_lock)
            {
                _connection ??= new DuckDBConnection("DataSource=:memory:");
                _connection.Open();
            }
        }
        return _connection;
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL query and return the first result")]
    public static object DuckQuery(string sql)
    {
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            var result = cmd.ExecuteScalar();
            return result ?? ExcelEmpty.Value;
        }
        catch (Exception ex)
        {
            return $"#ERROR: {ex.Message}";
        }
    }

    [ExcelFunction(Description = "Execute a DuckDB SQL query and return results as an array")]
    public static object[,] DuckQueryArray(string sql)
    {
        try
        {
            var conn = GetConnection();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            using var reader = cmd.ExecuteReader();

            var rows = new List<object[]>();
            var fieldCount = reader.FieldCount;

            // Add header row
            var headers = new object[fieldCount];
            for (int i = 0; i < fieldCount; i++)
            {
                headers[i] = reader.GetName(i);
            }
            rows.Add(headers);

            // Add data rows
            while (reader.Read())
            {
                var row = new object[fieldCount];
                for (int i = 0; i < fieldCount; i++)
                {
                    row[i] = reader.IsDBNull(i) ? ExcelEmpty.Value : reader.GetValue(i);
                }
                rows.Add(row);
            }

            if (rows.Count == 1)
            {
                // Only headers, no data
                return new object[1, fieldCount];
            }

            var result = new object[rows.Count, fieldCount];
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 0; j < fieldCount; j++)
                {
                    result[i, j] = rows[i][j];
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
    public static object DuckExecute(string sql)
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
}
