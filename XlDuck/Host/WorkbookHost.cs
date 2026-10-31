// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using XlDuck.Client;
using XlDuck.Core;
using XlDuck.Host;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlDuck;

internal readonly record struct CellAddress(nint Sheet, int Row, int Column);
internal sealed class CellBinding(string key, ResultLease lease, Task<object> value, bool previewOutput = false) : IDisposable
{
    internal string Key { get; set; } = key;
    internal ResultLease Lease { get; } = lease;
    internal Task<object> Value { get; } = value;
    internal bool PreviewOutput { get; } = previewOutput;
    internal bool Seen;
    internal bool Published;
    public void Dispose() => Lease.Dispose();
}
internal sealed class FormulaCell(string formula) : IDisposable
{
    internal string Formula = formula;
    internal bool Touched;
    internal ComputationFailure? ImmediateFailure;
    internal bool WaitingInputs;
    internal readonly Dictionary<string, CellBinding> Bindings = new(StringComparer.Ordinal);
    public void Dispose() { foreach (var binding in Bindings.Values) binding.Dispose(); Bindings.Clear(); }
}
internal sealed class OpenWorkbook(Excel.Workbook workbook)
{
    internal readonly Guid Id = Guid.NewGuid();
    internal readonly Excel.Workbook Workbook = workbook;
    internal string Name = workbook.Name;
    internal string Path = SavedPath(workbook);
    internal bool Date1904 = workbook.Date1904;
    internal WorkbookSettings Settings = new();
    internal WorkbookSession? Session;
    internal ExplorationSession? Exploration;
    internal ExplorationTransfer? ExplorationTransfer;
    internal int PreviewGeneration;
    internal bool SettingsSaved;
    internal bool Closed;
    internal long Revision;
    internal string? Status;
    internal bool HasMetadata;
    internal readonly Dictionary<CellAddress, FormulaCell> Cells = new();
    internal readonly Dictionary<nint, Excel.Worksheet> Sheets = new();
    internal static string SavedPath(Excel.Workbook workbook) => string.IsNullOrEmpty(workbook.Path) ? "" : workbook.FullName;
}

/// <summary>Excel main-thread owner of workbook identity, formula leases and publication.</summary>
internal static class WorkbookHost
{
    private static readonly Dictionary<nint, OpenWorkbook> Books = new();
    private static readonly Dictionary<nint, OpenWorkbook> Sheets = new();
    private static readonly List<Task> Disposals = [];
    private static readonly WorkbookSettingsStore SettingsStore = new(System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XlDuck", "connections"));
    private static readonly ConnectionCredentialBindings CredentialBindings = new(System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XlDuck", "credential-bindings"));
    private static Excel.Application? _application;
    private static System.Windows.Forms.Timer? _timer;
    private static volatile bool _stopped = true;
    internal static event Action? Changed;
    internal static IReadOnlyList<OpenWorkbook> OpenBooks => Books.Values.ToArray();
    internal static string InstallDirectory => System.IO.Path.GetDirectoryName(ExcelDnaUtil.XllPath)!;
    internal static Excel.Application Application => _application ?? throw new InvalidOperationException("XlDuck is not running.");
    private static nint Identity(object value)
    {
        var identity = Marshal.GetIUnknownForObject(value);
        Marshal.Release(identity);
        return identity;
    }

    internal static void Start()
    {
        _stopped = false;
        _application = (Excel.Application)ExcelDnaUtil.Application;
        _application.WorkbookOpen += Opened;
        ((Excel.AppEvents_Event)_application).NewWorkbook += Opened;
        _application.WorkbookAfterSave += Saved;
        _application.WorkbookBeforeClose += Closing;
        _application.WorkbookNewSheet += NewSheet;
        _application.SheetChange += Edited;
        _application.SheetCalculate += Calculated;
        _timer = new() { Interval = 750 };
        _timer.Tick += Tick;
        _timer.Start();
        Queue(Reconcile);
    }
    internal static void Queue(Action action)
    {
        if (_stopped) return;
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (_stopped) return;
            try { action(); } catch (Exception error) { Log.Error("Workbook callback", error); }
        });
    }
    private static void Tick(object? sender, EventArgs args) => Queue(Reconcile);
    private static void Opened(Excel.Workbook workbook) => Queue(Reconcile);
    private static void NewSheet(Excel.Workbook workbook, object sheet) => Queue(Reconcile);
    // BeforeClose precedes save prompts and cancellation. Only an absent workbook is disposed.
    private static void Closing(Excel.Workbook workbook, ref bool cancel) => Queue(Reconcile);
    private static void Saved(Excel.Workbook workbook, bool success)
    {
        if (!success) return;
        var identity = Identity(workbook);
        if (Books.TryGetValue(identity, out var saving))
        {
            saving.SettingsSaved = saving.HasMetadata && WorkbookMetadataAccess.SupportsFormat(workbook);
            if (saving.HasMetadata && !saving.SettingsSaved) saving.Status = "This file format cannot keep XlDuck connections. Save as .xlsx, .xlsm or .xlsb.";
            else if (saving.SettingsSaved) saving.Status = null;
        }
        Queue(() =>
        {
            Reconcile();
            Changed?.Invoke();
        });
    }
    internal static void Reconcile()
    {
        var live = new HashSet<nint>();
        foreach (Excel.Workbook workbook in Application.Workbooks)
        {
            if (workbook.Application.IsSandboxed) continue;
            var identity = Identity(workbook); live.Add(identity);
            if (!Books.TryGetValue(identity, out var book))
            {
                book = new(workbook); Books.Add(identity, book);
                _ = LoadSettingsAsync(book);
            }
            var dateChanged = book.Date1904 != workbook.Date1904;
            book.Name = workbook.Name; book.Path = OpenWorkbook.SavedPath(workbook); book.Date1904 = workbook.Date1904;
            var liveSheets = new HashSet<nint>();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                var name = $"[{workbook.Name}]{sheet.Name}";
                var reference = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, name);
                liveSheets.Add(reference.SheetId);
                book.Sheets[reference.SheetId] = sheet; Sheets[reference.SheetId] = book;
            }
            foreach (var removed in book.Sheets.Keys.Where(id => !liveSheets.Contains(id)).ToArray())
            {
                book.Sheets.Remove(removed); Sheets.Remove(removed);
                foreach (var address in book.Cells.Keys.Where(c => c.Sheet == removed).ToArray()) RemoveCell(book, address);
            }
            if (dateChanged) Recalculate(book);
        }
        foreach (var removed in Books.Keys.Where(key => !live.Contains(key)).ToArray())
        {
            var book = Books[removed]; book.Closed = true; book.Revision++;
            RetireSession(book); foreach (var sheet in book.Sheets.Keys) Sheets.Remove(sheet);
            book.Sheets.Clear(); Books.Remove(removed);
        }
        Disposals.RemoveAll(task => task.IsCompleted);
        Changed?.Invoke();
    }
    private static async Task LoadSettingsAsync(OpenWorkbook book)
    {
        var path = book.Path; var revision = book.Revision;
        try
        {
            // Capture COM metadata before any await. A present but invalid part must
            // never fall back to legacy settings or overwrite itself with defaults.
            var xml = WorkbookMetadataAccess.Read(book.Workbook);
            var settings = xml is null ? await SettingsStore.ReadForMigrationAsync(path).ConfigureAwait(false)
                : CredentialBindings.Resolve(WorkbookMetadata.Read(xml));
            Queue(() =>
            {
                if (book.Closed || book.Revision != revision) return;
                if (book.Path != path) { _ = LoadSettingsAsync(book); return; }
                try
                {
                    ApplySettings(book, settings ?? new(), save: xml is null && settings is not null);
                    if (xml is not null) { book.HasMetadata = true; book.SettingsSaved = true; }
                    else if (settings is not null) book.Status = "Connection settings migrated into this workbook. Save it to keep them with the file.";
                }
                catch (Exception error) { SettingsFailed(book, error); }
            });
        }
        catch (Exception error) { Queue(() => { if (!book.Closed && book.Revision == revision) SettingsFailed(book, error); }); }
    }
    private static void SettingsFailed(OpenWorkbook book, Exception error)
    {
        Log.Error("Read workbook connections", error);
        book.Status = "Could not load workbook connections: " + error.Message;
        Recalculate(book); Changed?.Invoke();
    }
    internal static OpenWorkbook Active()
    {
        var workbook = Application.ActiveWorkbook ?? throw new InvalidOperationException("Open a workbook first.");
        RequireOrdinaryWorkbook(workbook);
        if (!Books.ContainsKey(Identity(workbook))) Reconcile();
        return Books[Identity(workbook)];
    }
    internal static (OpenWorkbook Book, CellAddress Address, string Formula) Caller()
    {
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference reference)
            throw new ArgumentException("Use this function in a worksheet cell.");
        return Context(reference);
    }
    internal static (OpenWorkbook Book, CellAddress Address, string Formula) Context(ExcelReference reference)
    {
        if (reference.RowFirst != reference.RowLast || reference.ColumnFirst != reference.ColumnLast || reference.InnerReferences.Count != 1)
            throw new ArgumentException("Reference one worksheet cell.");
        if (!Sheets.TryGetValue(reference.SheetId, out var book))
        {
            Queue(Reconcile);
            throw new InvalidOperationException("Workbook is initializing or is in Protected View. XlDuck runs after Excel allows editing.");
        }
        RequireOrdinaryWorkbook(book.Workbook);
        var formula = Convert.ToString(XlCall.Excel(XlCall.xlfGetCell, 41, reference)) ?? "";
        return (book, new(reference.SheetId, reference.RowFirst, reference.ColumnFirst), formula);
    }
    internal static FormulaCell Touch(OpenWorkbook book, CellAddress address, string formula)
    {
        if (!book.Cells.TryGetValue(address, out var cell) || cell.Formula != formula)
        {
            cell?.Dispose(); book.Cells[address] = cell = new(formula);
        }
        if (!cell.Touched) { cell.ImmediateFailure = null; cell.WaitingInputs = false; }
        cell.Touched = true;
        return cell;
    }
    internal static void RequireOrdinaryWorkbook(Excel.Workbook workbook)
    {
        if (workbook.Application.IsSandboxed)
            throw new InvalidOperationException("XlDuck does not run in Protected View. Use Excel's Enable Editing only if you trust the workbook.");
    }
    internal static OpenWorkbook? FindSheet(Excel.Worksheet sheet)
    {
        var workbook = (Excel.Workbook)sheet.Parent;
        return workbook.Application.IsSandboxed ? null : Books.GetValueOrDefault(Identity(workbook));
    }
    internal static void ApplySettings(OpenWorkbook book, WorkbookSettings settings, bool save = true)
    {
        RequireOrdinaryWorkbook(book.Workbook);
        settings.Validate(); settings = settings.WithConnectionIds();
        if (save)
        {
            // These references come from the host dialog/command or legacy local
            // storage. Workbook XML is resolved separately and cannot supply one.
            foreach (var connection in settings.S3.Where(c => c.Credentials == XlDuck.Core.S3CredentialSource.AccessKey && c.CredentialId is not null))
                CredentialBindings.Save(connection);
            WorkbookMetadataAccess.Write(book.Workbook, settings);
            book.HasMetadata = true; book.SettingsSaved = false;
        }
        book.Revision++; book.Settings = settings; book.Status = null;
        Restart(book);
    }
    internal static void Restart(OpenWorkbook book)
    {
        RequireOrdinaryWorkbook(book.Workbook);
        if (book.Session is null && book.Status is not null) throw new InvalidOperationException(book.Status);
        book.Revision++;
        RetireSession(book);
        var worker = System.IO.Path.Combine(InstallDirectory, "worker", "XlDuck.Worker.exe");
        var settings = book.Settings;
        var readers = System.IO.Path.Combine(InstallDirectory, "readers");
        book.Session = new(id => new WorkerBackend(id, worker, settings.ForSession(readers, CredentialStore.Read)));
        book.Exploration = new(book.Session);
        book.ExplorationTransfer = new();
        Recalculate(book); Changed?.Invoke();
    }
    private static void RetireSession(OpenWorkbook book)
    {
        book.Exploration?.Dispose(); book.Exploration = null;
        book.ExplorationTransfer = null;
        foreach (var cell in book.Cells.Values) cell.Dispose(); book.Cells.Clear();
        if (book.Session is { } session)
        {
            book.Session = null;
            var disposal = session.DisposeAsync().AsTask(); Disposals.Add(disposal);
            _ = disposal.ContinueWith(task => { if (task.IsFaulted) Log.Write("Worker cleanup failed; restart XlDuck."); }, TaskScheduler.Default);
        }
    }
    internal static void Recalculate(OpenWorkbook book)
    {
        RequireOrdinaryWorkbook(book.Workbook);
        foreach (var sheet in book.Sheets.Values)
        {
            Excel.Range formulas;
            try { formulas = sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); }
            catch (COMException) { continue; }
            foreach (Excel.Range area in formulas.Areas)
            {
                var width = area.Columns.Count; var height = area.Rows.Count;
                var batch = Math.Max(1, 8192 / width);
                for (var offset = 0; offset < height; offset += batch)
                {
                    var top = area.Row + offset; var left = area.Column; var rows = Math.Min(batch, height - offset);
                    var block = sheet.Range[sheet.Cells[top, left], sheet.Cells[top + rows - 1, left + width - 1]];
                    object values = ((dynamic)block).Formula2;
                    for (var row = 0; row < rows; row++)
                    for (var column = 0; column < width; column++)
                    {
                        var formula = values is object[,] array ? Convert.ToString(array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)]) : Convert.ToString(values);
                        if (formula is null || !ExcelFormula.UsesXlDuck(formula)) continue;
                        var cell = (Excel.Range)sheet.Cells[top + row, left + column];
                        cell.Dirty();
                        if (Application.Calculation == Excel.XlCalculation.xlCalculationAutomatic) cell.Calculate();
                    }
                }
            }
        }
    }
    internal static void Publish(OpenWorkbook book, CellAddress address, FormulaCell cell, CellBinding binding) => Queue(() =>
    {
        if (book.Closed || !book.Cells.TryGetValue(address, out var current) || current != cell ||
            !cell.Bindings.TryGetValue(binding.Key, out var active) || active != binding) return;
        RequireOrdinaryWorkbook(book.Workbook);
        if (!book.Sheets.TryGetValue(address.Sheet, out var sheet)) return;
        var range = (Excel.Range)sheet.Cells[address.Row + 1, address.Column + 1];
        if (!string.Equals(Convert.ToString(((dynamic)range).Formula2), cell.Formula, StringComparison.Ordinal)) return;
        if (Application.Calculation == Excel.XlCalculation.xlCalculationAutomatic) { range.Dirty(); range.Calculate(); }
        Changed?.Invoke();
    });
    private static void Edited(object source, Excel.Range target)
    {
        if (source is not Excel.Worksheet sheet || FindSheet(sheet) is not { } book) return;
        foreach (var (address, cell) in book.Cells.ToArray())
        {
            if (!book.Sheets.TryGetValue(address.Sheet, out var owner) || Identity(owner) != Identity(sheet)) continue;
            var range = (Excel.Range)sheet.Cells[address.Row + 1, address.Column + 1];
            if (Application.Intersect(target, range) is null) continue;
            if (Convert.ToString(((dynamic)range).Formula2) != cell.Formula) RemoveCell(book, address);
        }
        // Excel can move a formula and its cached spill without invoking the UDF
        // at its new address. Re-establish ownership there through normal
        // evaluation; unchanged requests reuse the retained computation.
        Queue(() =>
        {
            if (book.Closed || FindSheet(sheet) != book) return;
            Excel.Range formulas;
            try { formulas = target.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); }
            catch (COMException) { return; }
            var changed = Application.Intersect(formulas, target);
            if (changed is null) return;
            var sheetId = book.Sheets.FirstOrDefault(pair => Identity(pair.Value) == Identity(sheet)).Key;
            foreach (Excel.Range range in changed.Cells)
            {
                var formula = Convert.ToString(((dynamic)range).Formula2) ?? "";
                if (!ExcelFormula.UsesXlDuck(formula)) continue;
                var address = new CellAddress(sheetId, range.Row - 1, range.Column - 1);
                if (book.Cells.TryGetValue(address, out var current) && current.Formula == formula) continue;
                range.Dirty();
                if (Application.Calculation == Excel.XlCalculation.xlCalculationAutomatic) range.Calculate();
            }
            Changed?.Invoke();
        });
        Changed?.Invoke();
    }
    private static void Calculated(object source)
    {
        if (source is not Excel.Worksheet sheet || FindSheet(sheet) is not { } book) return;
        foreach (var (address, cell) in book.Cells)
        {
            if (!cell.Touched || !book.Sheets.TryGetValue(address.Sheet, out var owner) || Identity(owner) != Identity(sheet)) continue;
            foreach (var key in cell.Bindings.Where(pair => !pair.Value.Seen).Select(pair => pair.Key).ToArray())
            {
                cell.Bindings[key].Dispose(); cell.Bindings.Remove(key);
            }
            foreach (var binding in cell.Bindings.Values) binding.Seen = false;
            cell.Touched = false;
        }
        Changed?.Invoke();
    }
    private static void RemoveCell(OpenWorkbook book, CellAddress address) { if (book.Cells.Remove(address, out var cell)) cell.Dispose(); }
    internal static void Stop()
    {
        if (_stopped) return;
        _stopped = true;
        _timer?.Stop(); _timer?.Dispose(); _timer = null;
        if (_application is not null)
        {
            try
            {
                _application.WorkbookOpen -= Opened; ((Excel.AppEvents_Event)_application).NewWorkbook -= Opened;
                _application.WorkbookAfterSave -= Saved; _application.WorkbookBeforeClose -= Closing;
                _application.WorkbookNewSheet -= NewSheet; _application.SheetChange -= Edited; _application.SheetCalculate -= Calculated;
            }
            catch (Exception error) { Log.Error("Workbook event cleanup", error); }
        }
        foreach (var book in Books.Values) { book.Closed = true; RetireSession(book); book.Sheets.Clear(); }
        Books.Clear(); Sheets.Clear(); _application = null;
        try { Task.WhenAll(Disposals).Wait(TimeSpan.FromSeconds(8)); } catch { }
        Disposals.Clear(); Changed = null;
    }
}
