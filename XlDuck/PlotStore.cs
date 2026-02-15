// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck;

/// <summary>
/// Represents a stored plot configuration.
/// </summary>
public class StoredPlot
{
    public string DataHandle { get; }
    public string Template { get; }
    public Dictionary<string, string> Overrides { get; }
    public DateTime CreatedUtc { get; }

    public StoredPlot(string dataHandle, string template, Dictionary<string, string> overrides)
    {
        DataHandle = dataHandle;
        Template = template;
        Overrides = overrides;
        CreatedUtc = DateTime.UtcNow;
    }
}

/// <summary>
/// Thread-safe store for plot configurations, keyed by handle.
/// </summary>
public static class PlotStore
{
    private static readonly Dictionary<string, StoredPlot> _plots = new();
    private static readonly Dictionary<string, int> _refCounts = new();
    private static readonly object _lock = new();
    private static long _nextId = 1;

    /// <summary>
    /// Store a plot configuration and return its handle.
    /// </summary>
    public static string Store(StoredPlot plot)
    {
        lock (_lock)
        {
            var id = _nextId++;
            var handle = $"duck://plot/{id}";
            _plots[handle] = plot;
            return handle;
        }
    }

    /// <summary>
    /// Retrieve a stored plot by handle.
    /// </summary>
    public static StoredPlot? Get(string handle)
    {
        lock (_lock)
        {
            return _plots.TryGetValue(handle, out var plot) ? plot : null;
        }
    }

    /// <summary>
    /// Check if a string is a plot handle.
    /// </summary>
    internal static bool IsHandle(string? value)
    {
        return value?.StartsWith("duck://plot/") == true;
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
            Log.Write($"[PlotStore] RefCount++ {handle}: {count + 1}");
        }
    }

    /// <summary>
    /// Decrement reference count for a handle. Removes plot when count reaches zero.
    /// </summary>
    internal static void DecrementRefCount(string handle)
    {
        lock (_lock)
        {
            if (_refCounts.TryGetValue(handle, out var count))
            {
                count--;
                Log.Write($"[PlotStore] RefCount-- {handle}: {count}");

                if (count <= 0)
                {
                    _refCounts.Remove(handle);
                    _plots.Remove(handle);
                    Log.Write($"[PlotStore] Evicted {handle}");
                }
                else
                {
                    _refCounts[handle] = count;
                }
            }
        }
    }
}
