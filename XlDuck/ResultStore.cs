// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck;

/// <summary>
/// Represents a stored query result as a DuckDB temp table.
/// </summary>
public class StoredResult
{
    public string DuckTableName { get; }
    public string[] ColumnNames { get; }
    public long RowCount { get; }

    public StoredResult(string duckTableName, string[] columnNames, long rowCount)
    {
        DuckTableName = duckTableName;
        ColumnNames = columnNames;
        RowCount = rowCount;
    }
}

/// <summary>
/// Thread-safe store for query results, keyed by handle.
/// </summary>
public static class ResultStore
{
    private static readonly Dictionary<string, StoredResult> _results = new();
    private static readonly Dictionary<string, int> _refCounts = new();
    private static readonly object _lock = new();
    private static long _nextId = 1;

    /// <summary>
    /// Store a result and return its handle (with dimensions).
    /// </summary>
    public static string Store(StoredResult result)
    {
        lock (_lock)
        {
            var id = _nextId++;
            var baseHandle = $"duck://table/{id}";
            _results[baseHandle] = result;
            // Return handle with dimensions: duck://table/123|10x4
            return $"{baseHandle}|{result.RowCount}x{result.ColumnNames.Length}";
        }
    }

    /// <summary>
    /// Retrieve a stored result by handle (strips dimension suffix if present).
    /// </summary>
    public static StoredResult? Get(string handle)
    {
        lock (_lock)
        {
            var baseHandle = GetBaseHandle(handle);
            return _results.TryGetValue(baseHandle, out var result) ? result : null;
        }
    }

    /// <summary>
    /// Strip dimension suffix from handle (duck://table/123|10x4 -> duck://table/123).
    /// </summary>
    internal static string GetBaseHandle(string handle)
    {
        var pipeIndex = handle.IndexOf('|');
        return pipeIndex >= 0 ? handle[..pipeIndex] : handle;
    }

    /// <summary>
    /// Check if a string is a valid handle format.
    /// </summary>
    internal static bool IsHandle(string value)
    {
        return value.StartsWith("duck://table/");
    }

    /// <summary>
    /// Increment reference count for a handle.
    /// </summary>
    internal static void IncrementRefCount(string handle)
    {
        lock (_lock)
        {
            var baseHandle = GetBaseHandle(handle);
            _refCounts.TryGetValue(baseHandle, out var count);
            _refCounts[baseHandle] = count + 1;
            System.Diagnostics.Debug.WriteLine($"[ResultStore] RefCount++ {baseHandle}: {count + 1}");
        }
    }

    /// <summary>
    /// Decrement reference count for a handle. Returns the evicted StoredResult if count reaches zero, null otherwise.
    /// Caller is responsible for dropping the DuckDB temp table.
    /// </summary>
    internal static StoredResult? DecrementRefCount(string handle)
    {
        lock (_lock)
        {
            var baseHandle = GetBaseHandle(handle);
            if (_refCounts.TryGetValue(baseHandle, out var count))
            {
                count--;
                System.Diagnostics.Debug.WriteLine($"[ResultStore] RefCount-- {baseHandle}: {count}");

                if (count <= 0)
                {
                    _refCounts.Remove(baseHandle);
                    _results.Remove(baseHandle, out var evicted);
                    System.Diagnostics.Debug.WriteLine($"[ResultStore] Evicted {baseHandle}");
                    return evicted;
                }
                else
                {
                    _refCounts[baseHandle] = count;
                }
            }
            return null;
        }
    }

    /// <summary>
    /// Get current reference count for a handle (for debugging).
    /// </summary>
    internal static int GetRefCount(string handle)
    {
        lock (_lock)
        {
            var baseHandle = GetBaseHandle(handle);
            return _refCounts.TryGetValue(baseHandle, out var count) ? count : 0;
        }
    }
}
