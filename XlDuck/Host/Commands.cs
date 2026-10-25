// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Text.Json;
using ExcelDna.Integration;
using XlDuck.Host;
using XlDuck.Core;

namespace XlDuck;

public static class DuckCommands
{
    private static OpenWorkbook Book(string name)
    {
        if (string.IsNullOrEmpty(name)) return WorkbookHost.Active();
        WorkbookHost.Reconcile();
        return WorkbookHost.OpenBooks.SingleOrDefault(book => string.Equals(book.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Workbook is not open in this Excel instance.");
    }
    [ExcelCommand(Name = "DuckRestartWorkbook")]
    public static void DuckRestartWorkbook(string workbookName) => WorkbookHost.Restart(Book(workbookName));
    [ExcelCommand(Name = "DuckSetConnections", Description = "Set non-secret connection settings as JSON and restart this workbook.")]
    public static void DuckSetConnections(string settingsJson, string workbookName)
    {
        if (settingsJson.Length > 64 * 1024) throw new ArgumentException("Connection settings are too large.");
        var settings = JsonSerializer.Deserialize<WorkbookSettings>(settingsJson) ?? throw new ArgumentException("Missing connection settings.");
        settings.Validate();
        var book = Book(workbookName);
        WorkbookHost.ApplySettings(book, settings);
    }
    [ExcelCommand(Name = "DuckSetResources", Description = "Set this workbook's resource budgets as JSON and restart its worker.")]
    public static void DuckSetResources(string resourcesJson, string workbookName)
    {
        if (resourcesJson.Length > 1024) throw new ArgumentException("Resource settings are too large.");
        var resources = JsonSerializer.Deserialize<ResourceSettings>(resourcesJson) ?? throw new ArgumentException("Missing resource settings.");
        resources.Validate(); var book = Book(workbookName);
        WorkbookHost.ApplySettings(book, book.Settings with { Resources = resources });
    }
    [ExcelCommand(Name = "DuckPauseQueries")]
    public static void DuckPauseQueries(string workbookName) => Book(workbookName).Session?.SetPaused(true);
    [ExcelCommand(Name = "DuckResumeQueries")]
    public static void DuckResumeQueries(string workbookName) => Book(workbookName).Session?.SetPaused(false);
    [ExcelCommand(Name = "DuckInterrupt")]
    public static void DuckInterrupt(string workbookName) => Book(workbookName).Session?.CancelPending();
    [ExcelCommand(Name = "DuckRetryFailed")]
    public static void DuckRetryFailed(string workbookName)
    {
        var book = Book(workbookName);
        book.Session?.InvalidateFailuresForRetry();
        book.PreviewGeneration++;
        foreach (var cell in book.Cells.Values)
        foreach (var key in cell.Bindings.Where(pair => pair.Value.Value.IsFaulted || pair.Value.Value.IsCanceled).Select(pair => pair.Key).ToArray())
        {
            cell.Bindings[key].Dispose(); cell.Bindings.Remove(key);
        }
        WorkbookHost.Recalculate(book);
    }
    [ExcelCommand(Name = "DuckTogglePreview")]
    public static void TogglePreview() => Preview.PreviewPaneManager.Instance.TogglePane();
}
