// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck;

/// <summary>
/// Represents a stored SQL fragment (deferred query).
/// </summary>
public class StoredFragment
{
    public string Sql { get; }
    public object[] Args { get; }

    public StoredFragment(string sql, object[] args)
    {
        Sql = sql;
        Args = args;
    }
}

/// <summary>
/// Thread-safe store for SQL fragments, keyed by handle.
/// </summary>
public static class FragmentStore
{
    private static readonly Dictionary<string, StoredFragment> _fragments = new();
    private static readonly Dictionary<string, int> _refCounts = new();
    private static readonly object _lock = new();
    private static long _nextId = 1;

    /// <summary>
    /// Store a fragment and return its handle.
    /// </summary>
    public static string Store(StoredFragment fragment)
    {
        lock (_lock)
        {
            var id = _nextId++;
            var handle = $"duck://f/{id}";
            _fragments[handle] = fragment;
            return handle;
        }
    }

    /// <summary>
    /// Retrieve a stored fragment by handle.
    /// </summary>
    public static StoredFragment? Get(string handle)
    {
        lock (_lock)
        {
            return _fragments.TryGetValue(handle, out var frag) ? frag : null;
        }
    }

    /// <summary>
    /// Check if a string is a fragment handle.
    /// </summary>
    internal static bool IsHandle(string? value)
    {
        return value?.StartsWith("duck://f/") == true;
    }

    /// <summary>
    /// Increment reference count for a handle.
    /// </summary>
    internal static void IncrementRefCount(string handle)
    {
        lock (_lock)
        {
            _refCounts.TryGetValue(handle, out var count);
            _refCounts[handle] = count + 1;
            System.Diagnostics.Debug.WriteLine($"[FragmentStore] RefCount++ {handle}: {count + 1}");
        }
    }

    /// <summary>
    /// Decrement reference count for a handle. Removes fragment when count reaches zero.
    /// </summary>
    internal static void DecrementRefCount(string handle)
    {
        lock (_lock)
        {
            if (_refCounts.TryGetValue(handle, out var count))
            {
                count--;
                System.Diagnostics.Debug.WriteLine($"[FragmentStore] RefCount-- {handle}: {count}");

                if (count <= 0)
                {
                    _refCounts.Remove(handle);
                    _fragments.Remove(handle);
                    System.Diagnostics.Debug.WriteLine($"[FragmentStore] Evicted {handle}");
                }
                else
                {
                    _refCounts[handle] = count;
                }
            }
        }
    }
}
