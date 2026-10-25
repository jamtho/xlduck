// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Runtime.InteropServices;
using System.Text.Json;
using ExcelDna.Integration;
using XlDuck.Client;
using XlDuck.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlDuck.TestAddIn;

// Integration harness only. Native Protected View enforcement, Save As and pane integration belong in the production host.
public sealed class AddIn : IExcelAddIn
{
    internal static AddIn? Instance;
    private Excel.Application? _application;
    public void AutoOpen()
    {
        Instance = this;
        _application = (Excel.Application)ExcelDnaUtil.Application;
        _application.SheetChange += OnSheetChange;
        _application.WorkbookBeforeClose += OnBeforeClose;
    }
    private static void OnBeforeClose(Excel.Workbook book, ref bool cancel) { if (!cancel) Functions.Close(book.Name); }
    private static void OnSheetChange(object sheetObject, Excel.Range target)
    {
        if (sheetObject is not Excel.Worksheet sheet) return;
        Excel.Workbook? book = null;
        try { book = (Excel.Workbook)sheet.Parent; Functions.Changed(book.Name, sheet.Name, target); }
        finally { if (book is not null) Marshal.ReleaseComObject(book); }
    }
    public void AutoClose()
    {
        if (_application is not null)
        {
            _application.SheetChange -= OnSheetChange;
            _application.WorkbookBeforeClose -= OnBeforeClose;
        }
        Functions.CloseAll();
        if (_application is not null) Marshal.ReleaseComObject(_application);
        _application = null; Instance = null;
        ExcelDna.Logging.LogDisplay.Hide();
        GC.Collect(); GC.WaitForPendingFinalizers();
    }
}

public static class Functions
{
    private sealed record Caller(string Book, string Sheet, int Row, int Column);
    private sealed class BookState(string workerPath, ConnectionSettings settings) : IDisposable
    {
        internal readonly string WorkerPath = workerPath;
        internal readonly ConnectionSettings Settings = settings;
        // Zero historical entries deliberately stress live ownership in the integration tests.
        internal readonly WorkbookSession Session = new(id => new WorkerBackend(id, workerPath, settings), new(HistoricalResults: 0));
        internal readonly Dictionary<Caller, Binding> Cells = new();
        public void Dispose()
        {
            foreach (var binding in Cells.Values) binding.Dispose();
            Cells.Clear();
            Session.DisposeAsync().AsTask().GetAwaiter().GetResult();
        }
    }
    private sealed class Binding(string key, ResultLease lease, Func<Task<object>> value) : IDisposable
    {
        internal readonly string Key = key;
        internal readonly ResultLease Lease = lease;
        internal readonly Task<object> Value = value();
        public void Dispose() => Lease.Dispose();
    }
    private static readonly Dictionary<string, BookState> Books = new();
    private static Caller GetCaller()
    {
        var reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        var name = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
        var close = name.LastIndexOf(']'); var open = name.LastIndexOf('[', close);
        return new(name[(open + 1)..close], name[(close + 1)..], reference.RowFirst, reference.ColumnFirst);
    }
    private static object Evaluate(string key, Func<BookState, Binding> create)
    {
        Caller? caller = null;
        BookState? state = null;
        try
        {
            caller = GetCaller();
            if (!Books.TryGetValue(caller.Book, out state)) return "#ERROR: Initialize this test workbook first.";
            if (!state.Cells.TryGetValue(caller, out var binding) || binding.Key != key)
            {
                var previous = binding;
                try { state.Cells[caller] = binding = create(state); }
                finally { previous?.Dispose(); }
            }
            if (binding.Value.IsCompletedSuccessfully) return binding.Value.Result;
            if (binding.Value.IsFaulted) return "#ERROR: " + binding.Value.Exception!.GetBaseException().Message;
            if (binding.Value.IsCanceled) return "#ERROR: Computation cancelled.";
            var current = binding;
            return ExcelAsyncUtil.Observe("XlDuckCoreTest", new object[] { state.Session.Id.ToString(), caller.Sheet, caller.Row, caller.Column, key, current.Lease.Handle.ToString() },
                () => new Observable(current.Value, () => Publish(caller, state, current)));
        }
        catch (Exception error)
        {
            if (caller is not null && state is not null && state.Cells.Remove(caller, out var stale)) stale.Dispose();
            return "#ERROR: " + error.Message;
        }
    }
    private static object EvaluateRequest(ComputationRequest request) => Evaluate(request.Identity(), state =>
    {
        var lease = state.Session.Acquire(request);
        return new(request.Identity(), lease, async () => (await lease.Completion.ConfigureAwait(false)).Handle.ToString());
    });
    private static object Query(string sql, bool fragment, params object[] supplied)
    {
        try
        {
            var args = supplied.ToList();
            while (args.Count > 0 && args[^1] is ExcelMissing or null) args.RemoveAt(args.Count - 1);
            var arguments = new List<QueryArgument>();
            foreach (var value in args)
            {
                if (value is ExcelError) return value;
                if (value is string error && error.StartsWith("#ERROR:", StringComparison.Ordinal)) return error;
                arguments.Add(value is string handle && handle.StartsWith("xlduck://", StringComparison.Ordinal)
                    ? QueryArgument.Reference(ResultHandle.Parse(handle))
                    : new QueryArgument(value is ExcelEmpty or ExcelMissing ? CellValue.Empty : CellValue.From(value)));
            }
            return EvaluateRequest(new(fragment ? RequestKind.Fragment : RequestKind.Query, sql, [.. arguments]));
        }
        catch (Exception error) { return "#ERROR: " + error.Message; }
    }
    [ExcelFunction]
    public static object CoreQuery(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!) => Query(sql, false, arg1, arg2, arg3);
    [ExcelFunction]
    public static object CoreFragment(string sql, object arg1 = null!, object arg2 = null!, object arg3 = null!) => Query(sql, true, arg1, arg2, arg3);
    [ExcelFunction]
    public static object CoreCapture(object[,] range)
    {
        try
        {
            var headers = Enumerable.Range(0, range.GetLength(1)).Select(column => Convert.ToString(range[0, column]) ?? "").ToImmutableArray();
            var rows = Enumerable.Range(1, range.GetLength(0) - 1).Select(row =>
                Enumerable.Range(0, range.GetLength(1)).Select(column => range[row, column] is ExcelEmpty ? CellValue.Empty : CellValue.From(range[row, column])).ToImmutableArray()).ToImmutableArray();
            return EvaluateRequest(ComputationRequest.CaptureRange(new(headers, rows)));
        }
        catch (Exception error) { return "#ERROR: " + error.Message; }
    }
    [ExcelFunction]
    public static object CorePlot(object input, string template)
    {
        if (input is ExcelError || input is string error && error.StartsWith("#ERROR:", StringComparison.Ordinal)) return input;
        try { return EvaluateRequest(ComputationRequest.Plot(ResultHandle.Parse(Convert.ToString(input)!), template)); }
        catch (Exception ex) { return "#ERROR: " + ex.Message; }
    }
    [ExcelFunction] public static object CoreScalar(object input) => Output(input, ReadPurpose.Scalar);
    [ExcelFunction] public static object CoreOut(object input) => Output(input, ReadPurpose.Output);
    [ExcelFunction] public static object CorePreview(object input) => Output(input, ReadPurpose.Preview);
    private static object Output(object input, ReadPurpose purpose)
    {
        if (input is ExcelError || input is string error && error.StartsWith("#ERROR:", StringComparison.Ordinal)) return input;
        try
        {
            var handle = ResultHandle.Parse(Convert.ToString(input)!);
            var key = purpose + ":" + handle;
            return Evaluate(key, state =>
            {
                var lease = state.Session.AcquireData(handle);
                return new(key, lease, async () =>
                {
                    await lease.Completion.ConfigureAwait(false);
                    var page = await state.Session.ReadAsync(lease.Handle, new(purpose, MaxRows: purpose == ReadPurpose.Preview ? 100 : 10000)).ConfigureAwait(false);
                    if (purpose == ReadPurpose.Preview) return JsonSerializer.Serialize(page);
                    if (purpose == ReadPurpose.Scalar) return Display(page.Rows[0][0]);
                    var values = new object[page.Rows.Length + 1, page.Columns.Length];
                    for (var column = 0; column < page.Columns.Length; column++)
                    {
                        values[0, column] = page.Columns[column].Name;
                        for (var row = 0; row < page.Rows.Length; row++) values[row + 1, column] = Display(page.Rows[row][column]);
                    }
                    return values;
                });
            });
        }
        catch (Exception exception) { return "#ERROR: " + exception.Message; }
    }
    private static object Display(CellValue value) => value.ToDisplayValue() ?? ExcelEmpty.Value;
    [ExcelCommand]
    public static void CoreInitialize(string workerPath, string workbookName)
    {
        Close(workbookName);
        Books[workbookName] = new(Path.GetFullPath(workerPath), new());
    }
    [ExcelCommand]
    public static void CoreInitializeWithSettings(string workerPath, string workbookName, string settingsJson)
    {
        var settings = JsonSerializer.Deserialize<ConnectionSettings>(settingsJson) ?? throw new ArgumentException("Missing settings.");
        settings.Validate();
        Close(workbookName);
        Books[workbookName] = new(Path.GetFullPath(workerPath), settings);
    }
    [ExcelFunction]
    public static string CoreStatistics(string workbookName) => Books.TryGetValue(workbookName, out var book) ? JsonSerializer.Serialize(book.Session.Statistics) : "";
    [ExcelCommand]
    public static void CoreRestart(string workbookName)
    {
        var path = Books[workbookName].WorkerPath;
        var settings = Books[workbookName].Settings;
        Close(workbookName);
        Books[workbookName] = new(path, settings);
        WithBook(workbookName, book =>
        {
            var sheets = book.Worksheets;
            try
            {
                for (var index = 1; index <= sheets.Count; index++)
                {
                    var sheet = (Excel.Worksheet)sheets[index]; var range = sheet.UsedRange;
                    try { range.Dirty(); sheet.Calculate(); }
                    finally { Marshal.ReleaseComObject(range); Marshal.ReleaseComObject(sheet); }
                }
            }
            finally { Marshal.ReleaseComObject(sheets); }
        });
    }
    [ExcelCommand] public static void CoreShutdown() => AddIn.Instance?.AutoClose();
    internal static void Close(string book) { if (Books.Remove(book, out var state)) state.Dispose(); }
    internal static void CloseAll() { foreach (var book in Books.Keys.ToArray()) Close(book); }
    internal static void Changed(string bookName, string sheetName, Excel.Range target)
    {
        if (!Books.TryGetValue(bookName, out var book)) return;
        foreach (var caller in book.Cells.Keys.Where(c => c.Sheet == sheetName).ToArray())
        {
            Excel.Range? cell = null, overlap = null;
            try
            {
                var sheet = target.Worksheet; var cells = sheet.Cells;
                try { cell = (Excel.Range)cells[caller.Row + 1, caller.Column + 1]; }
                finally { Marshal.ReleaseComObject(cells); Marshal.ReleaseComObject(sheet); }
                var application = (Excel.Application)ExcelDnaUtil.Application;
                overlap = application.Intersect(target, cell);
                if (overlap is null) continue;
                // Re-entering an unchanged formula must preserve its lease, even with no historical cache.
                // A changed Core formula replaces its binding when Excel evaluates it.
                if (((dynamic)cell).Formula2 is string formula && formula.Contains("Core", StringComparison.OrdinalIgnoreCase)) continue;
                if (book.Cells.Remove(caller, out var binding)) binding.Dispose();
            }
            finally { if (overlap is not null) Marshal.ReleaseComObject(overlap); if (cell is not null) Marshal.ReleaseComObject(cell); }
        }
    }
    private static void Publish(Caller caller, BookState state, Binding binding) => ExcelAsyncUtil.QueueAsMacro(() =>
    {
        if (state.Session.IsClosed || !state.Cells.TryGetValue(caller, out var current) || current != binding) return;
        var application = (Excel.Application)ExcelDnaUtil.Application;
        if (application.Calculation != Excel.XlCalculation.xlCalculationAutomatic) return;
        try
        {
            WithBook(caller.Book, book =>
            {
                var sheets = book.Worksheets; var sheet = (Excel.Worksheet)sheets[caller.Sheet]; var cells = sheet.Cells;
                var cell = (Excel.Range)cells[caller.Row + 1, caller.Column + 1];
                try { cell.Dirty(); cell.Calculate(); }
                finally { Marshal.ReleaseComObject(cell); Marshal.ReleaseComObject(cells); Marshal.ReleaseComObject(sheet); Marshal.ReleaseComObject(sheets); }
            });
        }
        catch (COMException) { /* The test workbook closed between completion and publication. */ }
    });
    private static void WithBook(string name, Action<Excel.Workbook> action)
    {
        var application = (Excel.Application)ExcelDnaUtil.Application; var books = application.Workbooks;
        Excel.Workbook? book = null;
        try { book = books[name]; action(book); }
        finally { if (book is not null) Marshal.ReleaseComObject(book); Marshal.ReleaseComObject(books); }
    }
    private sealed class Observable(Task<object> value, Action publish) : IExcelObservable
    {
        public IDisposable Subscribe(IExcelObserver observer)
        {
            var subscription = new Subscription();
            _ = Complete(observer, subscription);
            return subscription;
        }
        private async Task Complete(IExcelObserver observer, Subscription subscription)
        {
            object result;
            try { result = await value.ConfigureAwait(false); }
            catch (Exception error) { result = "#ERROR: " + error.GetBaseException().Message; }
            if (subscription.Disposed) return;
            observer.OnNext(result); observer.OnCompleted(); publish();
        }
        private sealed class Subscription : IDisposable
        {
            private int _disposed;
            public bool Disposed => Volatile.Read(ref _disposed) != 0;
            public void Dispose() => Interlocked.Exchange(ref _disposed, 1);
        }
    }
}
