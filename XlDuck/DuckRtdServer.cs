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
    // Time budget for synchronous execution before showing "Loading..."
    private const int TimeoutBudgetMs = 1000;

    // Track active topics and their associated handles
    private readonly ConcurrentDictionary<int, TopicInfo> _topics = new();

    private class TopicInfo
    {
        public Topic? Topic { get; set; }
        public string QueryType { get; set; } = "";
        public string? Handle { get; set; }
        public string Sql { get; set; } = "";
        public object[] Args { get; set; } = Array.Empty<object>();
        public bool IsComplete { get; set; }
        public ManualResetEventSlim? CompletionEvent { get; set; }
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
        System.Diagnostics.Debug.WriteLine($"[DuckRTD] ConnectData: TopicId={topic.TopicId}, IsReady={DuckFunctions.IsReady}, Info=[{string.Join(", ", topicInfo)}]");

        // topicInfo[0] = "query" or "frag"
        // topicInfo[1] = sql
        // topicInfo[2..] = serialized args (name1, value1, name2, value2, ...)

        var queryType = topicInfo[0];
        var sql = topicInfo[1];
        var args = topicInfo.Skip(2).Select(s => (object)s).ToArray();

        var completionEvent = new ManualResetEventSlim(false);
        var info = new TopicInfo
        {
            Topic = topic,
            QueryType = queryType,
            Sql = sql,
            Args = args,
            CompletionEvent = completionEvent
        };
        _topics[topic.TopicId] = info;

        // Check if query has @config sentinel OR depends on a blocked query - if so, wait for DuckConfigReady()
        bool requiresConfig = args.Any(a => a?.ToString() == DuckFunctions.ConfigSentinel);
        bool dependsOnBlocked = args.Any(a => a?.ToString()?.StartsWith(DuckFunctions.BlockedPrefix) == true);

        // Filter out sentinel from args before any execution path
        if (requiresConfig)
        {
            args = args.Where(a => a?.ToString() != DuckFunctions.ConfigSentinel).ToArray();
            info.Args = args;
        }

        if ((requiresConfig || dependsOnBlocked) && !DuckFunctions.IsReady)
        {
            System.Diagnostics.Debug.WriteLine($"[DuckRTD] @config sentinel found, waiting for DuckConfigReady()...");
            newValues = true;

            // Poll for ready flag in background, then execute query
            ThreadPool.QueueUserWorkItem(_ =>
            {
                while (!DuckFunctions.IsReady)
                {
                    Thread.Sleep(100);
                }
                System.Diagnostics.Debug.WriteLine($"[DuckRTD] DuckConfigReady() called, executing deferred query");
                ExecuteQuery(topic, info);
            });

            return DuckFunctions.ConfigBlockedStatus;
        }

        // Start query on background thread
        string? result = null;
        Exception? error = null;

        var queryThread = new Thread(() =>
        {
            try
            {
                if (queryType == "query")
                {
                    result = QueryExecutor.ExecuteQuery(sql, args);
                }
                else if (queryType == "frag")
                {
                    result = QueryExecutor.CreateFragment(sql, args);
                }
                else if (queryType == "plot")
                {
                    // For plot: sql is data handle, args[0] is template, rest are overrides
                    var template = args.Length > 0 ? args[0]?.ToString() ?? "" : "";
                    var overrideArgs = args.Length > 1 ? args.Skip(1).ToArray() : Array.Empty<object>();
                    result = QueryExecutor.CreatePlot(sql, template, overrideArgs);
                }
                else
                {
                    result = DuckFunctions.FormatError("internal", $"Unknown query type: {queryType}");
                }
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                completionEvent.Set();
            }
        });
        queryThread.IsBackground = true;
        queryThread.Start();

        // Wait for completion with timeout budget
        bool completedInTime = completionEvent.Wait(TimeoutBudgetMs);

        if (completedInTime)
        {
            // Query finished within budget - return result directly
            completionEvent.Dispose();
            info.CompletionEvent = null;

            if (error != null)
            {
                info.IsComplete = true;
                newValues = true;
                return DuckFunctions.FormatException(error);
            }

            info.Handle = result;
            info.IsComplete = true;

            // Increment refcount if we got a valid handle
            if (result != null && !DuckFunctions.IsErrorOrBlocked(result))
            {
                if (queryType == "query")
                    ResultStore.IncrementRefCount(result);
                else if (queryType == "frag")
                    FragmentStore.IncrementRefCount(result);
                else if (queryType == "plot")
                    PlotStore.IncrementRefCount(result);
            }

            newValues = true;
            System.Diagnostics.Debug.WriteLine($"[DuckRTD] Completed in budget: {result}");
            return result ?? DuckFunctions.FormatError("internal", "No result");
        }
        else
        {
            // Query still running - return placeholder and complete async
            System.Diagnostics.Debug.WriteLine($"[DuckRTD] Timeout, showing Loading...");

            // Continue waiting on another thread and update when done
            ThreadPool.QueueUserWorkItem(_ =>
            {
                completionEvent.Wait(); // Wait for completion
                completionEvent.Dispose();
                info.CompletionEvent = null;

                string finalResult;
                if (error != null)
                {
                    finalResult = DuckFunctions.FormatException(error);
                }
                else
                {
                    finalResult = result ?? DuckFunctions.FormatError("internal", "No result");
                }

                info.Handle = finalResult;
                info.IsComplete = true;

                // Increment refcount if we got a valid handle
                if (!DuckFunctions.IsErrorOrBlocked(finalResult))
                {
                    if (queryType == "query")
                        ResultStore.IncrementRefCount(finalResult);
                    else if (queryType == "frag")
                        FragmentStore.IncrementRefCount(finalResult);
                    else if (queryType == "plot")
                        PlotStore.IncrementRefCount(finalResult);
                }

                // Update the topic with the result
                System.Diagnostics.Debug.WriteLine($"[DuckRTD] Async complete: {finalResult}");
                topic.UpdateValue(finalResult);
            });

            newValues = true;
            return "Loading...";
        }
    }

    /// <summary>
    /// Execute a query asynchronously and update the topic when done.
    /// Used for deferred execution after DuckReady() is called.
    /// </summary>
    private void ExecuteQuery(Topic topic, TopicInfo info)
    {
        string? result = null;
        Exception? error = null;

        try
        {
            if (info.QueryType == "query")
            {
                result = QueryExecutor.ExecuteQuery(info.Sql, info.Args);
            }
            else if (info.QueryType == "frag")
            {
                result = QueryExecutor.CreateFragment(info.Sql, info.Args);
            }
            else if (info.QueryType == "plot")
            {
                // For plot: Sql is data handle, Args[0] is template, rest are overrides
                var template = info.Args.Length > 0 ? info.Args[0]?.ToString() ?? "" : "";
                var overrideArgs = info.Args.Length > 1 ? info.Args.Skip(1).ToArray() : Array.Empty<object>();
                result = QueryExecutor.CreatePlot(info.Sql, template, overrideArgs);
            }
            else
            {
                result = DuckFunctions.FormatError("internal", $"Unknown query type: {info.QueryType}");
            }
        }
        catch (Exception ex)
        {
            error = ex;
        }

        string finalResult;
        if (error != null)
        {
            finalResult = DuckFunctions.FormatException(error);
        }
        else
        {
            finalResult = result ?? DuckFunctions.FormatError("internal", "No result");
        }

        info.Handle = finalResult;
        info.IsComplete = true;

        // Increment refcount if we got a valid handle
        if (!DuckFunctions.IsErrorOrBlocked(finalResult))
        {
            if (info.QueryType == "query")
                ResultStore.IncrementRefCount(finalResult);
            else if (info.QueryType == "frag")
                FragmentStore.IncrementRefCount(finalResult);
            else if (info.QueryType == "plot")
                PlotStore.IncrementRefCount(finalResult);
        }

        System.Diagnostics.Debug.WriteLine($"[DuckRTD] Deferred complete: {finalResult}");
        topic.UpdateValue(finalResult);
    }

    protected override void DisconnectData(Topic topic)
    {
        System.Diagnostics.Debug.WriteLine($"[DuckRTD] DisconnectData: TopicId={topic.TopicId}");

        if (_topics.TryRemove(topic.TopicId, out var info))
        {
            // Decrement refcount if we had a valid handle
            if (info.Handle != null && !DuckFunctions.IsErrorOrBlocked(info.Handle))
            {
                if (ResultStore.IsHandle(info.Handle))
                {
                    var evicted = ResultStore.DecrementRefCount(info.Handle);
                    if (evicted != null)
                    {
                        // Drop the DuckDB temp table now that it's no longer referenced
                        DuckFunctions.DropTempTable(evicted.DuckTableName);
                    }
                }
                else if (FragmentStore.IsHandle(info.Handle))
                    FragmentStore.DecrementRefCount(info.Handle);
                else if (PlotStore.IsHandle(info.Handle))
                    PlotStore.DecrementRefCount(info.Handle);
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

    public static string CreatePlot(string dataHandle, string template, object[] args)
    {
        return DuckFunctions.CreatePlotInternal(dataHandle, template, args);
    }
}
