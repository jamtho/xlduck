// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlDuck.Preview;

internal sealed class PreviewPaneManager : IDisposable
{
    private sealed record PaneEntry(Guid Book, CustomTaskPane Pane, PreviewController Controller);
    private static PreviewPaneManager? _instance;
    internal static PreviewPaneManager Instance => _instance ??= new();
    private readonly Dictionary<int, PaneEntry> _panes = new();
    private PreviewPaneManager()
    {
        WorkbookHost.Changed += Refresh;
        WorkbookHost.Application.SheetSelectionChange += Selected;
    }
    private void Selected(object sheet, Excel.Range target) => Refresh();
    internal void TogglePane()
    {
        var book = WorkbookHost.Active();
        var window = WorkbookHost.Application.ActiveWindow;
        if (window is null) return;
        if (_panes.TryGetValue(window.Hwnd, out var existing)) existing.Pane.Visible = !existing.Pane.Visible;
        else
        {
            var pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(PreviewPane), "XlDuck Preview");
            pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight; pane.Width = 650;
            _panes.Add(window.Hwnd, new(book.Id, pane, new((PreviewPane)pane.ContentControl)));
            pane.Visible = true;
        }
        Refresh();
    }
    internal bool IsPaneVisible() => WorkbookHost.Application.ActiveWindow is { } window && _panes.TryGetValue(window.Hwnd, out var pane) && pane.Pane.Visible;
    private void Refresh()
    {
        try
        {
            var live = new HashSet<int>();
            foreach (var book in WorkbookHost.OpenBooks)
            foreach (Excel.Window window in book.Workbook.Windows)
            {
                live.Add(window.Hwnd);
                if (!_panes.TryGetValue(window.Hwnd, out var pane) || !pane.Pane.Visible) continue;
                var range = window.RangeSelection as Excel.Range;
                if (range is null || Convert.ToDouble((object)range.CountLarge) != 1) { pane.Controller.Clear(); continue; }
                // Excel owns spill geometry. Follow its parent rather than infer
                // a rectangle from old results, which may have shrunk or moved.
                if ((object)((dynamic)range).HasSpill is true) range = (Excel.Range)((dynamic)range).SpillParent;
                var sheetId = book.Sheets.FirstOrDefault(pair => pair.Value.Name == range.Worksheet.Name).Key;
                if (sheetId == 0) continue;
                var address = new CellAddress(sheetId, range.Row - 1, range.Column - 1);
                object? value = range.Value2;
                if (book.Cells.TryGetValue(address, out var formula) && formula.Formula == Convert.ToString(((dynamic)range).Formula2))
                {
                    var sources = formula!.Bindings.Values.Where(binding => binding.PreviewOutput && (!formula.Touched || binding.Seen))
                        .Select(binding => binding.Lease.CellText).Distinct(StringComparer.Ordinal).ToArray();
                    // A formula combining multiple outputs has no single source.
                    if (sources.Length == 1 && formula.ImmediateFailure is null && !formula.WaitingInputs &&
                        value is not System.Runtime.InteropServices.ErrorWrapper && !(value is int code && code < 0))
                        value = sources[0];
                }
                pane.Controller.Select(book, address, value);
            }
            foreach (var id in _panes.Keys.Where(id => !live.Contains(id)).ToArray()) Remove(id);
        }
        catch (System.Runtime.InteropServices.COMException) { /* Excel may be closing a window. The next reconciliation retries. */ }
    }
    private void Remove(int id)
    {
        var pane = _panes[id]; _panes.Remove(id); pane.Controller.Dispose();
        try { pane.Pane.Delete(); } catch (System.Runtime.InteropServices.COMException) { }
    }
    public void Dispose()
    {
        WorkbookHost.Changed -= Refresh;
        try { WorkbookHost.Application.SheetSelectionChange -= Selected; }
        catch (Exception error) { Log.Error("Preview event cleanup", error); }
        foreach (var id in _panes.Keys.ToArray()) Remove(id);
    }
    internal static void CloseAll() { _instance?.Dispose(); _instance = null; }
}
