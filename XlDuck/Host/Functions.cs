// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Globalization;
using System.Text.Json;
using ExcelDna.Integration;
using XlDuck.Core;
using XlDuck.Host;

namespace XlDuck;

public sealed class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        WorkbookHost.Start();
        XlCall.Excel(XlCall.xlcOnKey, "^+P", "DuckTogglePreview");
    }
    public void AutoClose()
    {
        Log.Write("Add-in shutdown started.");
        try { Preview.PreviewPaneManager.CloseAll(); }
        catch (Exception error) { Log.Error("Preview shutdown", error); }
        finally { WorkbookHost.Stop(); }
        try { XlCall.Excel(XlCall.xlcOnKey, "^+P"); } catch { }
        Log.Write("Add-in shutdown completed.");
    }
}

public static class DuckFunctions
{
    internal static string Error(Exception error) => "#ERROR: " + error.GetBaseException().Message;
    private static object Evaluate(string key, Func<OpenWorkbook, CellBinding> create)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "XlDuck";
        try
        {
            var (book, address, formula) = WorkbookHost.Caller();
            key += ":date1904=" + book.Date1904;
            if (!book.Enabled || book.Session is null) return "#BLOCKED: Enable XlDuck for this workbook from the XlDuck ribbon.";
            if (!book.Cells.TryGetValue(address, out var cell) || cell.Formula != formula)
            {
                cell?.Dispose(); book.Cells[address] = cell = new(formula);
            }
            cell.Touched = true;
            if (!cell.Bindings.TryGetValue(key, out var binding))
            {
                binding = create(book); binding.Key = key; cell.Bindings.Add(key, binding);
                var current = binding; var currentCell = cell;
                _ = binding.Value.ContinueWith(_ => WorkbookHost.Publish(book, address, currentCell, current), TaskScheduler.Default);
            }
            binding.Seen = true;
            if (binding.Value.IsCompletedSuccessfully) return binding.Value.Result;
            if (binding.Value.IsFaulted) return Error(binding.Value.Exception!);
            if (binding.Value.IsCanceled) return "#ERROR: Computation cancelled. Use Retry failed results or restart this workbook.";
            // RTD is a fallback notification. Targeted publication reads the completed cached value,
            // including arrays, without putting it through RTD's value and throttle limitations.
            _ = ExcelAsyncUtil.Observe("XlDuck", new object[] { book.Id.ToString(), book.Session.Id.ToString(),
                address.Sheet.ToInt64(), address.Row, address.Column, key, binding.Lease.Handle.ToString() }, () => new CompletionSignal(binding.Value));
            return ExcelError.ExcelErrorNA;
        }
        catch (Exception error) { return Error(error); }
    }
    private static object Request(ComputationRequest request, ReadPurpose? output = null)
    {
        var key = request.Identity() + ":" + output;
        return Evaluate(key, book =>
        {
            var lease = book.Session!.Acquire(request);
            return new(key, lease, output is { } purpose ? Read(book.Session, lease, purpose, book.Date1904, 10000)
                : Handle(lease));
        });
    }
    private static async Task<object> Handle(ResultLease lease) => (await lease.Completion.ConfigureAwait(false)).Handle.ToString();
    private static QueryArgument Argument(object? value)
    {
        if (value is object[,] array)
        {
            if (array.Length != 1) throw new ArgumentException("Use DuckCapture for a range containing more than one cell.");
            value = array[0, 0];
        }
        if (value is ExcelError) throw new ArgumentException("An input contains an Excel error.");
        return value is string handle && handle.StartsWith("xlduck://", StringComparison.Ordinal)
            ? QueryArgument.Reference(ResultHandle.Parse(handle))
            : new QueryArgument(value is ExcelEmpty or ExcelMissing ? CellValue.Empty : CellValue.From(value));
    }
    private static object Query(string sql, RequestKind kind, ReadPurpose? output, params object[] supplied)
    {
        try
        {
            var args = supplied.ToList();
            while (args.Count > 0 && args[^1] is ExcelMissing or null) args.RemoveAt(args.Count - 1);
            foreach (var value in args)
                if (value is ExcelError || value is string error && (error.StartsWith("#ERROR:", StringComparison.Ordinal) || error.StartsWith("#BLOCKED:", StringComparison.Ordinal))) return value;
            return Request(new(kind, sql, [.. args.Select(Argument)]), output);
        }
        catch (Exception error) { return Error(error); }
    }
    [ExcelFunction(IsMacroType = true, Description = "Read and compute SQL. Bind values or handles with positional ? arguments.")]
    public static object DuckQuery(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!, object arg4 = null!, object arg5 = null!, object arg6 = null!, object arg7 = null!, object arg8 = null!) =>
        Query(sql, RequestKind.Query, null, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
    [ExcelFunction(IsMacroType = true, Description = "Create a checked SQL recipe for composition. Use DuckQuery to retain a source snapshot.")]
    public static object DuckFrag(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!, object arg4 = null!, object arg5 = null!, object arg6 = null!, object arg7 = null!, object arg8 = null!) =>
        Query(sql, RequestKind.Fragment, null, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
    [ExcelFunction(IsMacroType = true, Description = "Read and compute SQL, then spill its bounded result with column headers.")]
    public static object DuckQueryOut(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!, object arg4 = null!, object arg5 = null!, object arg6 = null!, object arg7 = null!, object arg8 = null!) =>
        Query(sql, RequestKind.Query, ReadPurpose.Output, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
    [ExcelFunction(IsMacroType = true, Description = "Read and compute SQL returning exactly one row and one column.")]
    public static object DuckQueryOutScalar(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!, object arg4 = null!, object arg5 = null!, object arg6 = null!, object arg7 = null!, object arg8 = null!) =>
        Query(sql, RequestKind.Query, ReadPurpose.Scalar, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
    [ExcelFunction(IsMacroType = true, Description = "Capture a range with a header row, preserving its cell types.")]
    public static object DuckCapture(object[,] range)
    {
        try
        {
            if (range.Length == 0) throw new ArgumentException("Capture needs a header row.");
            if (range.Cast<object>().Any(value => value is ExcelError)) throw new ArgumentException("Capture contains an Excel error. Correct the source cells first.");
            var headers = Enumerable.Range(0, range.GetLength(1)).Select(column =>
                range[0, column] is ExcelEmpty ? "" : Convert.ToString(range[0, column], CultureInfo.InvariantCulture) ?? "").ToImmutableArray();
            var rows = Enumerable.Range(1, range.GetLength(0) - 1).Select(row => Enumerable.Range(0, range.GetLength(1)).Select(column =>
                range[row, column] is ExcelEmpty ? CellValue.Empty : CellValue.From(range[row, column])).ToImmutableArray()).ToImmutableArray();
            return Request(ComputationRequest.CaptureRange(new(headers, rows)));
        }
        catch (Exception error) { return Error(error); }
    }
    [ExcelFunction(IsMacroType = true, Description = "Spill a retained table or fragment with headers. Oversized results fail without truncation.")]
    public static object DuckOut(object handle, int maxRows = 10000) => Output(handle, ReadPurpose.Output, maxRows);
    [ExcelFunction(IsMacroType = true, Description = "Return exactly one row and one column from a table or fragment.")]
    public static object DuckScalar(object handle) => Output(handle, ReadPurpose.Scalar, 1);
    private static object Output(object input, ReadPurpose purpose, int maxRows)
    {
        if (input is ExcelError || input is string failure && failure.StartsWith('#')) return input;
        try
        {
            if (maxRows < 1 || maxRows > 1000000) throw new ArgumentException("maxRows must be between 1 and 1000000.");
            var handle = ResultHandle.Parse(Convert.ToString(input)!); var key = $"{purpose}:{maxRows}:{handle}";
            return Evaluate(key, book =>
            {
                var lease = book.Session!.AcquireData(handle);
                return new(key, lease, Read(book.Session, lease, purpose, book.Date1904, maxRows));
            });
        }
        catch (Exception error) { return Error(error); }
    }
    private static async Task<object> Read(WorkbookSession session, ResultLease lease, ReadPurpose purpose, bool date1904, int maxRows)
    {
        await lease.Completion.ConfigureAwait(false);
        var page = await session.ReadAsync(lease.Handle, new(purpose, maxRows)).ConfigureAwait(false);
        object Display(CellValue value) => ExcelValues.Display(value, date1904) ?? ExcelEmpty.Value;
        if (purpose == ReadPurpose.Scalar) return Display(page.Rows[0][0]);
        var output = new object[page.Rows.Length + 1, page.Columns.Length];
        for (var column = 0; column < page.Columns.Length; column++)
        {
            output[0, column] = page.Columns[column].Name;
            for (var row = 0; row < page.Rows.Length; row++) output[row + 1, column] = Display(page.Rows[row][column]);
        }
        return output;
    }
    [ExcelFunction(IsMacroType = true, Description = "Create a plot handle. Select its cell to view the chart in the preview pane.")]
    public static object DuckPlot(object handle, string template, object name1 = null!, object value1 = null!, object name2 = null!, object value2 = null!, object name3 = null!, object value3 = null!, object name4 = null!, object value4 = null!, object name5 = null!, object value5 = null!, object name6 = null!, object value6 = null!)
    {
        if (handle is ExcelError || handle is string failure && failure.StartsWith('#')) return handle;
        try
        {
            template = template.ToLowerInvariant();
            if (template is not ("bar" or "line" or "point" or "area" or "histogram" or "heatmap" or "boxplot" or "map")) throw new ArgumentException("Unknown chart template.");
            var args = new object[] { name1, value1, name2, value2, name3, value3, name4, value4, name5, value5, name6, value6 };
            var options = new Dictionary<string, string>(StringComparer.Ordinal);
            for (var i = 0; i < args.Length; i += 2)
            {
                if (args[i] is null or ExcelMissing) continue;
                var name = Convert.ToString(args[i], CultureInfo.InvariantCulture)!.ToLowerInvariant();
                if (name is not ("x" or "y" or "color" or "label" or "tooltip" or "title" or "value" or "xmin" or "xmax" or "ymin" or "ymax" or "lat" or "lon" or "basemap")) throw new ArgumentException("Unknown plot option: " + name);
                if (args[i + 1] is null or ExcelMissing or ExcelError) throw new ArgumentException("A plot option needs a value.");
                if (!options.TryAdd(name, Convert.ToString(args[i + 1], CultureInfo.InvariantCulture)!)) throw new ArgumentException("Duplicate plot option: " + name);
            }
            return Request(ComputationRequest.Plot(ResultHandle.Parse(Convert.ToString(handle)!), template, options));
        }
        catch (Exception error) { return Error(error); }
    }
    [ExcelFunction(IsMacroType = true, Description = "Inspect a result: schema (spilled column names/types), rows, or definition (JSON recipe).")]
    public static object DuckInspect(object input, string field = "schema")
    {
        if (input is ExcelError || input is string failure && failure.StartsWith('#')) return input;
        try
        {
            var handle = ResultHandle.Parse(Convert.ToString(input)!); field = field.ToLowerInvariant();
            var key = "inspect:" + field + ":" + handle;
            return Evaluate(key, book =>
            {
                var definition = book.Session!.GetDefinition(handle);
                var date1904 = book.Date1904;
                var lease = field == "definition" ? book.Session.AcquireExisting(handle) : book.Session.AcquireData(handle);
                async Task<object> Inspect()
                {
                    var result = await lease.Completion.ConfigureAwait(false);
                    if (field == "rows") return ExcelValues.Display(CellValue.From(result.RowCount), date1904)!;
                    if (field == "definition")
                    {
                        var json = JsonSerializer.Serialize(definition);
                        return json.Length <= 30000 ? json : throw new ArgumentException("Definition exceeds one cell. Inspect its source formula.");
                    }
                    if (field != "schema") throw new ArgumentException("Choose schema, rows or definition.");
                    var schema = new object[result.Columns.Length + 1, 2]; schema[0, 0] = "Column"; schema[0, 1] = "Type";
                    for (var i = 0; i < result.Columns.Length; i++) { schema[i + 1, 0] = result.Columns[i].Name; schema[i + 1, 1] = result.Columns[i].Type; }
                    return schema;
                }
                return new(key, lease, Inspect());
            });
        }
        catch (Exception error) { return Error(error); }
    }
    [ExcelFunction(IsMacroType = true)] public static object DuckDate(double serial) => Date(serial, false);
    [ExcelFunction(IsMacroType = true)] public static object DuckDateTime(double serial) => Date(serial, true);
    private static object Date(double serial, bool time) { try { return ExcelValues.DateText(serial, WorkbookHost.Caller().Book.Date1904, time); } catch (Exception error) { return Error(error); } }
    [ExcelFunction] public static string DuckVersion() => "0.2.0-dev";
    [ExcelFunction] public static string DuckLibraryVersion() => "v1.5.6";
    [ExcelFunction(IsMacroType = true, IsVolatile = true, Description = "Workbook enablement and computation diagnostics as JSON; no SQL or credentials.")]
    public static string DuckStatus()
    {
        try { var book = WorkbookHost.Caller().Book; return JsonSerializer.Serialize(new { enabled = book.Enabled, remembered = book.Remembered, paused = book.Session?.IsPaused ?? false, session = book.Session?.Id, runtime = Environment.Version.ToString(), statistics = book.Session?.Statistics, message = book.Status }); }
        catch (Exception error) { return Error(error); }
    }
    private sealed class CompletionSignal(Task task) : IExcelObservable
    {
        public IDisposable Subscribe(IExcelObserver observer)
        {
            var subscription = new Subscription();
            _ = task.ContinueWith(_ => { if (!subscription.Disposed) { observer.OnNext(1d); observer.OnCompleted(); } }, TaskScheduler.Default);
            return subscription;
        }
        private sealed class Subscription : IDisposable
        {
            private int _disposed;
            internal bool Disposed => Volatile.Read(ref _disposed) != 0;
            public void Dispose() => Interlocked.Exchange(ref _disposed, 1);
        }
    }
}
