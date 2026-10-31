// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Core;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace XlDuck.Preview;

internal sealed class PreviewController(PreviewPane pane) : IDisposable
{
    private CancellationTokenSource? _pending;
    private string? _key;
    private string? _cell;
    private bool _hasSuccess;
    private bool _disposed;
    internal void Select(OpenWorkbook book, CellAddress address, object? value)
    {
        var location = book.Id + ":" + address;
        var sameCell = _cell == location;
        var pending = book.Cells.TryGetValue(address, out var formula) && formula.Bindings.Values.Any(binding => !binding.Value.IsCompleted);
        var text = Convert.ToString(value);
        var key = $"{location}:{book.Session?.Id}:{book.PreviewGeneration}:{pending}:{book.Session?.IsPaused}:{text}";
        if (_key == key) return;
        _key = key; _cell = location;
        if (!sameCell) _hasSuccess = false;
        _pending?.Cancel(); _pending?.Dispose(); _pending = new();
        if (book.Session is null) { Status("Initializing XlDuck for this workbook…"); return; }
        if (pending) { Status((book.Session!.IsPaused ? "Paused" : "Computing") + (_hasSuccess ? " — showing the previous result." : "…")); return; }
        // A dependent formula can propagate #N/A before it calls an XlDuck function.
        // Value2 exposes Excel errors as COM HRESULT integers, not Excel-DNA enum values.
        if (value is ExcelError || value is ErrorWrapper || value is int code && code < 0)
        {
            Status("Excel reports an unavailable input or error." + (_hasSuccess ? " Showing the previous result." : "")); return;
        }
        if (text?.StartsWith("#ERROR:", StringComparison.Ordinal) == true) { Status(text); return; }
        if (!CellHandle.IsToken(text))
        {
            _hasSuccess = false;
            pane.SetState(new EmptyPreviewModel { Title = "Preview", Message = "Select a handle or a cell in a DuckOut / DuckQueryOut result." }); return;
        }
        Status("Loading preview" + (_hasSuccess ? " — showing the previous result." : "…"));
        _ = LoadAsync(book.Session, text, _pending.Token);
    }
    private void Status(string message) => pane.SetState(new StatusPreviewModel { Title = "Preview", Message = message, RetainPrevious = _hasSuccess });
    internal void Clear()
    {
        _key = null; _cell = null; _hasSuccess = false;
        _pending?.Cancel();
        pane.SetState(new EmptyPreviewModel { Title = "Preview", Message = "Select one handle cell or a cell in a DuckOut / DuckQueryOut result." });
    }
    private async Task LoadAsync(WorkbookSession session, string value, CancellationToken token)
    {
        try
        {
            var model = await PreviewDataProvider.GetPreviewAsync(session, value, token).ConfigureAwait(false);
            WorkbookHost.Queue(() => { if (!_disposed && !token.IsCancellationRequested) { _hasSuccess = true; pane.SetState(model); } });
        }
        catch (OperationCanceledException) { }
        catch (Exception error) { WorkbookHost.Queue(() => { if (!_disposed && !token.IsCancellationRequested) Status(DuckFunctions.Error(error)); }); }
    }
    public void Dispose() { _disposed = true; _pending?.Cancel(); _pending?.Dispose(); }
}
