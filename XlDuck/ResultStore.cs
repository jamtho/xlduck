namespace XlDuck;

/// <summary>
/// Represents a stored query result in memory.
/// </summary>
public class StoredResult
{
    public string[] ColumnNames { get; }
    public Type[] ColumnTypes { get; }
    public List<object?[]> Rows { get; }

    public StoredResult(string[] columnNames, Type[] columnTypes, List<object?[]> rows)
    {
        ColumnNames = columnNames;
        ColumnTypes = columnTypes;
        Rows = rows;
    }
}

/// <summary>
/// Thread-safe store for query results, keyed by handle.
/// </summary>
public static class ResultStore
{
    private static readonly Dictionary<string, StoredResult> _results = new();
    private static readonly object _lock = new();
    private static long _nextId = 1;

    /// <summary>
    /// Store a result and return its handle.
    /// </summary>
    public static string Store(StoredResult result)
    {
        lock (_lock)
        {
            var id = _nextId++;
            var handle = $"duck://t/{id}";
            _results[handle] = result;
            return handle;
        }
    }

    /// <summary>
    /// Retrieve a stored result by handle.
    /// </summary>
    public static StoredResult? Get(string handle)
    {
        lock (_lock)
        {
            return _results.TryGetValue(handle, out var result) ? result : null;
        }
    }

    /// <summary>
    /// Check if a string is a valid handle format.
    /// </summary>
    public static bool IsHandle(string value)
    {
        return value.StartsWith("duck://t/");
    }
}
