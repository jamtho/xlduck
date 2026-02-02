// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using ExcelDna.Integration.Rtd;
using System.Collections.Concurrent;

namespace XlDuck;

/// <summary>
/// RTD server for DuckQuery functions. Provides lifecycle tracking for handle reference counting.
/// </summary>
public class DuckRtdServer : ExcelRtdServer
{
    // Track active topics and their associated handles
    private readonly ConcurrentDictionary<int, TopicInfo> _topics = new();

    private class TopicInfo
    {
        public string? Handle { get; set; }
        public string Sql { get; set; } = "";
        public object[] Args { get; set; } = Array.Empty<object>();
        public bool IsComplete { get; set; }
    }

    protected override bool ServerStart()
    {
        System.Diagnostics.Debug.WriteLine("[DuckRTD] Server started");
        return true;
    }

    protected override void ServerTerminate()
    {
        System.Diagnostics.Debug.WriteLine("[DuckRTD] Server terminated");
    }

    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        System.Diagnostics.Debug.WriteLine($"[DuckRTD] ConnectData: TopicId={topic.TopicId}, Info=[{string.Join(", ", topicInfo)}]");

        // topicInfo[0] = "query" or "frag"
        // topicInfo[1] = sql
        // topicInfo[2..] = serialized args (name1, value1, name2, value2, ...)

        var queryType = topicInfo[0];
        var sql = topicInfo[1];
        var args = topicInfo.Skip(2).Select(s => (object)s).ToArray();

        var info = new TopicInfo { Sql = sql, Args = args };
        _topics[topic.TopicId] = info;

        // Execute query synchronously for now (we'll add timeout budget later)
        try
        {
            string result;
            if (queryType == "query")
            {
                result = QueryExecutor.ExecuteQuery(sql, args);
            }
            else if (queryType == "frag")
            {
                result = QueryExecutor.CreateFragment(sql, args);
            }
            else
            {
                result = $"#ERROR: Unknown query type: {queryType}";
            }

            info.Handle = result;
            info.IsComplete = true;

            // Increment refcount if we got a valid handle
            if (!result.StartsWith("#ERROR"))
            {
                if (queryType == "query")
                    ResultStore.IncrementRefCount(result);
                else if (queryType == "frag")
                    FragmentStore.IncrementRefCount(result);
            }

            newValues = true;
            return result;
        }
        catch (Exception ex)
        {
            info.IsComplete = true;
            newValues = true;
            return $"#ERROR: {ex.Message}";
        }
    }

    protected override void DisconnectData(Topic topic)
    {
        System.Diagnostics.Debug.WriteLine($"[DuckRTD] DisconnectData: TopicId={topic.TopicId}");

        if (_topics.TryRemove(topic.TopicId, out var info))
        {
            // Decrement refcount if we had a valid handle
            if (info.Handle != null && !info.Handle.StartsWith("#ERROR"))
            {
                if (ResultStore.IsHandle(info.Handle))
                    ResultStore.DecrementRefCount(info.Handle);
                else if (FragmentStore.IsHandle(info.Handle))
                    FragmentStore.DecrementRefCount(info.Handle);
            }
        }
    }
}

/// <summary>
/// Executes queries outside the main DuckFunctions class for use by RTD server.
/// </summary>
public static class QueryExecutor
{
    public static string ExecuteQuery(string sql, object[] args)
    {
        return DuckFunctions.ExecuteQueryInternal(sql, args);
    }

    public static string CreateFragment(string sql, object[] args)
    {
        return DuckFunctions.CreateFragmentInternal(sql, args);
    }
}
