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
        public int Epoch { get; set; }
        public ManualResetEventSlim? CompletionEvent { get; set; }
        public CancellationTokenSource? Cts { get; set; }
    }

    protected override bool ServerStart()
    {
        Log.Write("[DuckRTD] Server started");
        return true;
    }

    protected override void ServerTerminate()
    {
        Log.Write("[DuckRTD] Server terminated");
    }

    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        var connectSw = System.Diagnostics.Stopwatch.StartNew();

        // topicInfo[0] = query type: "query", "query-config", "frag", "frag-config", "plot", "capture"
        // topicInfo[1] = sql (or content hash for capture)
        // topicInfo[2..] = serialized positional args (value1, value2, ...)

        var rawQueryType = topicInfo[0];
        bool requiresConfig = rawQueryType.EndsWith("-config");
        var queryType = requiresConfig ? rawQueryType.Replace("-config", "") : rawQueryType;
        var sql = topicInfo[1];
        var args = topicInfo.Skip(2).Select(s => (object)s).ToArray();

        Log.Write($"\u2500\u2500\u2500 TopicId={topic.TopicId} [{queryType}] {new string('\u2500', 40)}");
        Log.Write($"    {sql}");
        if (args.Length > 0)
            Log.Write($"    args: {string.Join(", ", args)}");

        var epoch = DuckFunctions.InterruptEpoch;
        var completionEvent = new ManualResetEventSlim(false);
        var info = new TopicInfo
        {
            Topic = topic,
            QueryType = queryType,
            Sql = sql,
            Args = args,
            Epoch = epoch,
            CompletionEvent = completionEvent
        };
        _topics[topic.TopicId] = info;

        // Check if query requires config OR depends on a config-blocked query - if so, wait for DuckConfigReady()
        bool dependsOnBlocked = args.Any(a => a?.ToString() == DuckFunctions.ConfigBlockedStatus);

        if ((requiresConfig || dependsOnBlocked) && !DuckFunctions.IsReady)
        {
            Log.Write($"[DuckRTD] Config required, waiting for DuckConfigReady()...");
            newValues = true;

            // Poll for ready flag in background, then execute query
            ThreadPool.QueueUserWorkItem(_ =>
            {
                while (!DuckFunctions.IsReady)
                {
                    Thread.Sleep(100);
                }
                Log.Write($"[DuckRTD] DuckConfigReady() called, executing deferred query");
                DuckFunctions.SetThreadEpoch(info.Epoch);
                DuckFunctions.SetThreadTopicId(topic.TopicId);
                ExecuteQuery(topic, info);
            });

            return DuckFunctions.ConfigBlockedStatus;
        }

        // Check if queries are paused - defer execution until resumed
        if (DuckFunctions.QueriesPaused)
        {
            newValues = true;
            SpawnDeferredThread(topic, info);
            return DuckFunctions.PausedBlockedStatus;
        }

        // Start query on background thread
        string? result = null;
        Exception? error = null;
        string? resolvedSql = null;

        var queryThread = new Thread(() =>
        {
            DuckFunctions.SetThreadEpoch(epoch);
            DuckFunctions.SetThreadTopicId(topic.TopicId);
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
                else if (queryType == "capture")
                {
                    // For capture: sql holds the content hash
                    result = QueryExecutor.CaptureRange(sql);
                }
                else
                {
                    result = DuckFunctions.FormatError("internal", $"Unknown query type: {queryType}");
                }
            }
            catch (Exception ex)
            {
                error = ex;
                // Capture resolved SQL from this thread before it exits
                resolvedSql = DuckFunctions.ConsumeThreadResolvedSql();
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
                if (error is OperationCanceledException && DuckFunctions.QueriesPaused)
                {
                    newValues = true;
                    SpawnDeferredThread(topic, info);
                    return DuckFunctions.PausedBlockedStatus;
                }
                info.IsComplete = true;
                newValues = true;
                DuckFunctions.SetThreadResolvedSql(resolvedSql);
                var errorResult = DuckFunctions.FormatException(error);
                Log.Write($"[{topic.TopicId}] -> {errorResult} ({connectSw.ElapsedMilliseconds}ms)");
                return errorResult;
            }

            info.Handle = result;
            info.IsComplete = true;

            // Increment refcount if we got a valid handle
            if (result != null && !DuckFunctions.IsErrorOrBlocked(result))
            {
                if (queryType == "query" || queryType == "capture")
                    ResultStore.IncrementRefCount(result);
                else if (queryType == "frag")
                    FragmentStore.IncrementRefCount(result);
                else if (queryType == "plot")
                    PlotStore.IncrementRefCount(result);
            }

            newValues = true;
            Log.Write($"[{topic.TopicId}] -> {result} ({connectSw.ElapsedMilliseconds}ms)");
            return result ?? DuckFunctions.FormatError("internal", "No result");
        }
        else
        {
            // Query still running - return placeholder and complete async
            Log.Write($"[{topic.TopicId}] timeout -> Loading...");

            // Continue waiting on another thread and update when done
            ThreadPool.QueueUserWorkItem(_ =>
            {
                completionEvent.Wait(); // Wait for completion
                completionEvent.Dispose();
                info.CompletionEvent = null;

                if (error is OperationCanceledException && DuckFunctions.QueriesPaused)
                {
                    topic.UpdateValue(DuckFunctions.PausedBlockedStatus);
                    SpawnDeferredThread(topic, info);
                    return;
                }

                string finalResult;
                if (error != null)
                {
                    DuckFunctions.SetThreadResolvedSql(resolvedSql);
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
                    if (queryType == "query" || queryType == "capture")
                        ResultStore.IncrementRefCount(finalResult);
                    else if (queryType == "frag")
                        FragmentStore.IncrementRefCount(finalResult);
                    else if (queryType == "plot")
                        PlotStore.IncrementRefCount(finalResult);
                }

                // Update the topic with the result
                Log.Write($"[{topic.TopicId}] -> {finalResult}");
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
            if (info.Epoch != DuckFunctions.InterruptEpoch)
                throw new OperationCanceledException("Query cancelled");

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
            else if (info.QueryType == "capture")
            {
                result = QueryExecutor.CaptureRange(info.Sql);
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

        if (error is OperationCanceledException && DuckFunctions.QueriesPaused)
        {
            topic.UpdateValue(DuckFunctions.PausedBlockedStatus);
            SpawnDeferredThread(topic, info);
            return;
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
            if (info.QueryType == "query" || info.QueryType == "capture")
                ResultStore.IncrementRefCount(finalResult);
            else if (info.QueryType == "frag")
                FragmentStore.IncrementRefCount(finalResult);
            else if (info.QueryType == "plot")
                PlotStore.IncrementRefCount(finalResult);
        }

        Log.Write($"[{topic.TopicId}] -> {finalResult}");
        topic.UpdateValue(finalResult);
    }

    private void SpawnDeferredThread(Topic topic, TopicInfo info)
    {
        info.Cts ??= new CancellationTokenSource();
        var ct = info.Cts.Token;
        ThreadPool.QueueUserWorkItem(_ =>
        {
            if (!DuckFunctions.WaitForUnpause(ct)) return;
            if (ct.IsCancellationRequested) return;
            // Args containing blocked/error status will fail; skip and let
            // Excel recalculate with correct args when dependencies resolve.
            if (info.Args.Any(a => DuckFunctions.IsErrorOrBlocked(a?.ToString())))
                return;
            info.Epoch = DuckFunctions.InterruptEpoch;
            DuckFunctions.SetThreadEpoch(info.Epoch);
            DuckFunctions.SetThreadTopicId(topic.TopicId);
            ExecuteQuery(topic, info);
        });
    }

    protected override void DisconnectData(Topic topic)
    {
        Log.Write($"[{topic.TopicId}] disconnect");

        if (_topics.TryRemove(topic.TopicId, out var info))
        {
            info.Cts?.Cancel();
            // Decrement refcount if we had a valid handle
            if (info.Handle != null && !DuckFunctions.IsErrorOrBlocked(info.Handle))
            {
                if (ResultStore.IsHandle(info.Handle))
                {
                    var evicted = ResultStore.DecrementRefCount(info.Handle);
                    if (evicted != null)
                    {
                        var tableName = evicted.DuckTableName;
                        ThreadPool.QueueUserWorkItem(_ => DuckFunctions.DropTempTable(tableName));
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

    public static string CaptureRange(string hash)
    {
        var data = DuckFunctions.TakePendingCapture(hash);
        if (data == null)
            return DuckFunctions.FormatError("internal", "Capture data not found (hash expired)");
        return DuckFunctions.CaptureRangeInternal(data);
    }
}
