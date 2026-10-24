// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlDuck.ExecutionPrototype;

namespace XlDuck.PrototypeAddIn;

public sealed class AddIn : IExcelAddIn
{
    internal static AddIn? Instance { get; private set; }
    private Microsoft.Office.Interop.Excel.Application? _application;
    public void AutoOpen()
    {
        Instance = this;
        ExcelIntegration.RegisterUnhandledExceptionHandler(error =>
        {
            File.AppendAllText(Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!, "prototype-diagnostics.log"),
                DateTime.UtcNow.ToString("O") + " " + error + Environment.NewLine);
            return "#ERROR: Prototype exception; see prototype-diagnostics.log.";
        });
        _application = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        _application.WorkbookBeforeClose += OnWorkbookClosing;
    }
    private void OnWorkbookClosing(Microsoft.Office.Interop.Excel.Workbook workbook, ref bool cancel)
    {
        if (!cancel) Functions.CloseWorkbook(workbook.Name);
    }
    public void AutoClose()
    {
        try
        {
            if (_application is not null) _application.WorkbookBeforeClose -= OnWorkbookClosing;
            Functions.CloseAll();
        }
        finally
        {
            Functions.SaveDiagnostics();
            ExcelDna.Logging.LogDisplay.Hide();
            if (_application is not null) System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);
            _application = null;
            Instance = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}

public static class Functions
{
    // Workbook names are sufficient for this controlled prototype. Production must survive Save As.
    private static readonly Dictionary<(string Book, string Mode), Session> Sessions = new();
    private static readonly Dictionary<string, string> Publishers = new();
    private static string? _workerPath;
    private static string WorkerPath => _workerPath ?? Environment.GetEnvironmentVariable("XLDUCK_PROTOTYPE_WORKER")
        ?? throw new InvalidOperationException("XLDUCK_PROTOTYPE_WORKER is not set.");

    private sealed record Caller(string Book, string Sheet, int Row, int Column);
    private static Caller GetCaller()
    {
        var reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        var name = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
        var close = name.LastIndexOf(']');
        var open = name.LastIndexOf('[', close);
        return new(name[(open + 1)..close], name[(close + 1)..], reference.RowFirst, reference.ColumnFirst);
    }
    private static Session GetSession(string book, string mode)
    {
        if (mode is not ("inproc" or "worker")) throw new ArgumentException("Mode must be inproc or worker.");
        var key = (book, mode);
        if (!Sessions.TryGetValue(key, out var session))
        {
            var worker = mode == "worker" ? WorkerPath : "";
            Sessions[key] = session = new Session(() => mode == "worker" ? new WorkerBackend(worker) : new LocalBackend());
        }
        return session;
    }
    private static object Evaluate(string mode, Request request, Func<Result, object> convert)
    {
        try
        {
            var caller = GetCaller();
            var session = GetSession(caller.Book, mode);
            if (session.TryGetCompleted(request, out var completed))
                return completed!.Error is null ? convert(completed) : "#ERROR: " + completed.Error;
            var key = JsonSerializer.Serialize(request with { Id = Guid.Empty });
            var targeted = Publishers.GetValueOrDefault(caller.Book) == "targeted";
            return ExcelAsyncUtil.Observe("XlDuckPrototype", new object[] { session.Id, key, caller.Sheet, caller.Row, caller.Column },
                () => new Computation(session, request, convert,
                    targeted ? () => Publish(caller, session) : null));
        }
        catch (Exception ex) { return "#ERROR: " + ex.Message; }
    }

    [ExcelFunction(Description = "EXPERIMENT ONLY: materialize SQL using __input and optional $value.")]
    public static object ProtoQuery(string mode, string sql, object input, object value = null!)
    {
        try
        {
            if (input is ExcelError) return input;
            if (input is string error && error.StartsWith("#ERROR:")) return error;
            var handle = input is ExcelEmpty or ExcelMissing or null ? "" : Convert.ToString(input)!;
            var atom = value is ExcelEmpty or ExcelMissing or null ? null : Atom.From(value);
            return Evaluate(mode, new(Guid.NewGuid(), "query", sql, handle, atom), result => result.Handle);
        }
        catch (Exception ex) { return "#ERROR: " + ex.Message; }
    }

    [ExcelFunction]
    public static object ProtoScalar(string mode, object input)
    {
        if (input is ExcelError || input is string error && error.StartsWith("#ERROR:")) return input;
        return Evaluate(mode, new(Guid.NewGuid(), "rows", Input: Convert.ToString(input)!, Limit: 1),
            result => result.Rows is { Length: > 0 } ? ExcelValue(result.Rows[0][0]) : ExcelEmpty.Value);
    }

    [ExcelFunction]
    public static object ProtoOut(string mode, object input)
    {
        if (input is ExcelError || input is string error && error.StartsWith("#ERROR:")) return input;
        return Evaluate(mode, new(Guid.NewGuid(), "rows", Input: Convert.ToString(input)!, Limit: 10000), result =>
        {
            var output = new object[result.Rows!.Length + 1, result.Columns!.Length];
            for (var column = 0; column < result.Columns.Length; column++)
            {
                output[0, column] = result.Columns[column];
                for (var row = 0; row < result.Rows.Length; row++) output[row + 1, column] = ExcelValue(result.Rows[row][column]);
            }
            return output;
        });
    }

    [ExcelFunction]
    public static object ProtoPreview(string mode, object input)
    {
        if (input is ExcelError || input is string error && error.StartsWith("#ERROR:")) return input;
        return Evaluate(mode, new(Guid.NewGuid(), "rows", Input: Convert.ToString(input)!, Limit: 20),
            result => JsonSerializer.Serialize(new { result.Columns, result.Rows }));
    }

    [ExcelFunction]
    public static object ProtoCapture(string mode, object[,] data)
    {
        try
        {
            if (data.GetLength(1) != 1) return "#ERROR: Prototype captures one numeric column.";
            var rows = Enumerable.Range(0, data.GetLength(0)).Select(row => new[] { Atom.From(data[row, 0]) }).ToArray();
            return Evaluate(mode, new(Guid.NewGuid(), "capture", Capture: rows), result => result.Handle);
        }
        catch (Exception ex) { return "#ERROR: " + ex.Message; }
    }

    [ExcelCommand]
    public static void ProtoSetWorker(string path) => _workerPath = Path.GetFullPath(path);

    [ExcelCommand]
    public static void ProtoSetPublisher(string publisher, string workbookName)
    {
        if (publisher is not ("rtd" or "targeted")) throw new ArgumentException("Unknown publisher.");
        Publishers[workbookName] = publisher;
    }

    [ExcelCommand]
    public static void ProtoRestart(string workbookName)
    {
        dynamic application = ExcelDnaUtil.Application;
        string book = workbookName;
        foreach (var key in Sessions.Keys.Where(key => key.Book == book).ToArray())
        {
            var session = Sessions[key];
            Sessions.Remove(key);
            _ = session.DisposeAsync();
        }
        // Explicit command: recalculate only this workbook's sheets, not other workbooks.
        dynamic workbooks = application.Workbooks;
        dynamic workbook = workbooks[book];
        dynamic sheets = workbook.Worksheets;
        try
        {
            for (var index = 1; index <= (int)sheets.Count; index++)
            {
                dynamic sheet = sheets[index];
                dynamic range = sheet.UsedRange;
                try { range.Dirty(); sheet.Calculate(); }
                finally { Marshal.ReleaseComObject(range); Marshal.ReleaseComObject(sheet); }
            }
        }
        finally { Marshal.ReleaseComObject(sheets); Marshal.ReleaseComObject(workbook); Marshal.ReleaseComObject(workbooks); }
    }

    [ExcelFunction]
    public static double ProtoExecutions(string mode, object dependency) => GetSession(GetCaller().Book, mode).Executions;

    [ExcelFunction]
    public static string ProtoSessionInfo(string workbookName, string mode) =>
        Sessions.TryGetValue((workbookName, mode), out var session)
            ? JsonSerializer.Serialize(new { session.Id, session.Executions, session.Pending, session.WorkerProcessId }) : "";

    [ExcelCommand]
    public static void ProtoCrashWorker(string workbookName) => Sessions[(workbookName, "worker")].CrashWorker();

    [ExcelFunction]
    public static string ProtoDiagnostics() => (string)typeof(ExcelDna.Logging.LogDisplay)
        .GetMethod("GetAllText", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!
        .Invoke(null, null)!;

    [ExcelCommand]
    public static void ProtoShutdown() => AddIn.Instance?.AutoClose();

    internal static void SaveDiagnostics() => File.AppendAllText(
        Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!, "prototype-diagnostics.log"), ProtoDiagnostics());

    internal static void CloseAll()
    {
        foreach (var session in Sessions.Values) session.DisposeAsync().AsTask().GetAwaiter().GetResult();
        Sessions.Clear();
    }
    internal static void CloseWorkbook(string book)
    {
        foreach (var key in Sessions.Keys.Where(key => key.Book == book).ToArray())
        {
            var session = Sessions[key];
            Sessions.Remove(key);
            _ = session.DisposeAsync();
        }
        Publishers.Remove(book);
    }
    private static object ExcelValue(Atom atom) => atom.Kind == "null" ? ExcelEmpty.Value : atom.ToObject();

    private static void Publish(Caller caller, Session session) => ExcelAsyncUtil.QueueAsMacro(() =>
    {
        if (session.IsClosed) return;
        try
        {
            dynamic application = ExcelDnaUtil.Application;
            if ((int)application.Calculation != -4105) return; // Never override manual calculation.
            dynamic books = application.Workbooks;
            dynamic? book = null, sheets = null, sheet = null, cells = null, range = null;
            try
            {
                book = books[caller.Book]; sheets = book.Worksheets; sheet = sheets[caller.Sheet];
                cells = sheet.Cells; range = cells[caller.Row + 1, caller.Column + 1];
                // Excel evaluates the current formula; no result is written over a later edit.
                range.Dirty(); range.Calculate();
            }
            finally
            {
                foreach (object? item in new object?[] { range, cells, sheet, sheets, book, books })
                    if (item is not null) Marshal.ReleaseComObject(item);
            }
        }
        catch (System.Runtime.InteropServices.COMException) { /* Workbook or cell no longer available. */ }
    });

    private sealed class Computation(Session session, Request request, Func<Result, object> convert, Action? publish) : IExcelObservable
    {
        public IDisposable Subscribe(IExcelObserver observer)
        {
            var cancellation = new CancellationTokenSource();
            _ = RunAsync(observer, cancellation.Token);
            return new Subscription(cancellation);
        }
        private async Task RunAsync(IExcelObserver observer, CancellationToken token)
        {
            try
            {
                var result = await session.GetAsync(request, token).ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                observer.OnNext(convert(result));
                observer.OnCompleted();
                publish?.Invoke();
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                observer.OnNext("#ERROR: " + ex.Message);
                observer.OnCompleted();
                publish?.Invoke();
            }
        }
        private sealed class Subscription(CancellationTokenSource source) : IDisposable
        {
            public void Dispose() { source.Cancel(); source.Dispose(); }
        }
    }
}
