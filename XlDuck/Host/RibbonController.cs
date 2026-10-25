// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;

namespace XlDuck;

[ComVisible(true)]
[Guid("A1B2C3D4-E5F6-4A5B-9C8D-7E6F5A4B3C2D")]
public sealed class RibbonController : ExcelRibbon
{
    private IRibbonUI? _ribbon;
    public override void OnBeginShutdown(ref Array custom)
    {
        new AddIn().AutoClose();
        base.OnBeginShutdown(ref custom);
    }
    public override void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        _ribbon = null;
        new AddIn().AutoClose();
        base.OnDisconnection(removeMode, ref custom);
        GC.Collect(); GC.WaitForPendingFinalizers();
    }
    public override string GetCustomUI(string ribbonId)
    {
        using var stream = typeof(RibbonController).Assembly.GetManifestResourceStream("XlDuck.Ribbon.xml")!;
        using var reader = new StreamReader(stream); return reader.ReadToEnd();
    }
    public void OnRibbonLoad(IRibbonUI ribbon) { _ribbon = ribbon; WorkbookHost.Changed += () => _ribbon?.Invalidate(); }
    private static void Run(Action action) { try { action(); } catch (Exception error) { MessageBox.Show(error.Message, "XlDuck", MessageBoxButtons.OK, MessageBoxIcon.Information); } }
    public void OnConnections(IRibbonControl control) => Run(() => { using var dialog = new ConnectionsDialog(WorkbookHost.Active()); dialog.ShowDialog(); });
    public void OnResources(IRibbonControl control) => Run(() =>
    {
        var book = WorkbookHost.Active();
        using var dialog = new ResourcesDialog(book.Name, book.Settings.Resources);
        if (dialog.ShowDialog() == DialogResult.OK) WorkbookHost.ApplySettings(book, book.Settings with { Resources = dialog.Settings });
    });
    public void OnRestart(IRibbonControl control) => Run(() => DuckCommands.DuckRestartWorkbook(""));
    public void OnRetry(IRibbonControl control) => Run(() => DuckCommands.DuckRetryFailed(""));
    public void OnCancelQuery(IRibbonControl control) => Run(() => DuckCommands.DuckInterrupt(""));
    public void OnQueryPauseToggle(IRibbonControl control, bool pressed) => Run(() => { if (pressed) DuckCommands.DuckPauseQueries(""); else DuckCommands.DuckResumeQueries(""); });
    public bool GetQueryPaused(IRibbonControl control) { try { return WorkbookHost.Active().Session?.IsPaused ?? false; } catch { return false; } }
    public void OnPreviewPaneToggle(IRibbonControl control, bool pressed) => Run(DuckCommands.TogglePreview);
    public bool GetPreviewPanePressed(IRibbonControl control) { try { return Preview.PreviewPaneManager.Instance.IsPaneVisible(); } catch { return false; } }
    public void OnVersionClick(IRibbonControl control) => MessageBox.Show($"XlDuck {DuckFunctions.DuckVersion()}\nDuckDB {DuckFunctions.DuckLibraryVersion()}\n64-bit Excel · one worker per workbook", "XlDuck");
}
