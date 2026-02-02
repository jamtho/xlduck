// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

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
    internal static bool IsHandle(string value)
    {
        return value.StartsWith("duck://t/");
    }
}
