// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using XlDuck.Host;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlDuck;

/// <summary>Excel main-thread access to the one XlDuck custom XML part.</summary>
internal static class WorkbookMetadataAccess
{
    internal static bool SupportsFormat(Excel.Workbook workbook) => workbook.FileFormat is
        Excel.XlFileFormat.xlOpenXMLWorkbook or Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled or Excel.XlFileFormat.xlExcel12 or
        Excel.XlFileFormat.xlOpenXMLTemplate or Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled or Excel.XlFileFormat.xlOpenXMLStrictWorkbook;
    internal static string? Read(Excel.Workbook workbook)
    {
        WorkbookHost.RequireOrdinaryWorkbook(workbook);
        var parts = workbook.CustomXMLParts.SelectByNamespace(WorkbookMetadata.Namespace);
        if (parts.Count > 1) throw new InvalidOperationException("The workbook contains duplicate XlDuck settings. Open Connections and apply the intended settings to repair them.");
        return parts.Count == 0 ? null : parts[1].XML;
    }
    internal static void Write(Excel.Workbook workbook, WorkbookSettings settings)
    {
        WorkbookHost.RequireOrdinaryWorkbook(workbook);
        if (!SupportsFormat(workbook)) throw new InvalidOperationException("Save this workbook as .xlsx, .xlsm or .xlsb before applying XlDuck settings. Its current file format cannot retain connection metadata.");
        var xml = WorkbookMetadata.Write(settings);
        var parts = workbook.CustomXMLParts.SelectByNamespace(WorkbookMetadata.Namespace);
        if (parts.Count == 1 && parts[1].XML == xml) return;
        // Add before removing the old part so a failed Add preserves existing settings.
        var previous = Enumerable.Range(1, parts.Count).Select(index => parts[index]).ToArray();
        workbook.CustomXMLParts.Add(xml);
        foreach (var part in previous) part.Delete();
        workbook.Saved = false;
    }
}
