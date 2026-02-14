// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using XlDuck.Preview;

namespace XlDuck;

/// <summary>
/// Ribbon controller for XlDuck tab.
/// </summary>
[ComVisible(true)]
[Guid("A1B2C3D4-E5F6-4A5B-9C8D-7E6F5A4B3C2D")]
public class RibbonController : ExcelRibbon
{
    private IRibbonUI? _ribbon;

    public override string GetCustomUI(string ribbonId)
    {
        Log.Write($"[Ribbon] GetCustomUI called for ribbonId={ribbonId}");
        var xml = GetEmbeddedRibbonXml();
        Log.Write($"[Ribbon] Returning XML ({xml.Length} chars)");
        return xml;
    }

    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        Log.Write("[Ribbon] OnRibbonLoad called - ribbon loaded successfully");
    }

    public void OnPreviewPaneToggle(IRibbonControl control, bool pressed)
    {
        Log.Write($"[Ribbon] OnPreviewPaneToggle clicked, pressed={pressed}");
        try
        {
            PreviewPaneManager.Instance.TogglePane();
            _ribbon?.InvalidateControl("PreviewPaneToggle");
            Log.Write("[Ribbon] TogglePane completed");
        }
        catch (Exception ex)
        {
            Log.Error("OnPreviewPaneToggle", ex);
        }
    }

    public bool GetPreviewPanePressed(IRibbonControl control)
    {
        Log.Write("[Ribbon] GetPreviewPanePressed called");
        try
        {
            var pressed = PreviewPaneManager.Instance.IsPaneVisible();
            Log.Write($"[Ribbon] GetPreviewPanePressed returning {pressed}");
            return pressed;
        }
        catch (Exception ex)
        {
            Log.Error("GetPreviewPanePressed", ex);
            return false;
        }
    }

    public void OnCancelQuery(IRibbonControl control)
    {
        try
        {
            DuckFunctions.Interrupt();
        }
        catch (Exception ex)
        {
            Log.Error("OnCancelQuery", ex);
        }
    }

    public void OnVersionClick(IRibbonControl control)
    {
        try
        {
            var addInVersion = DuckFunctions.DuckVersion();
            var libVersion = DuckFunctions.DuckLibraryVersion();

            MessageBox.Show(
                $"XLDuck Add-in: v{addInVersion}\nDuckDB Library: {libVersion}",
                "XLDuck Version",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Error getting version: {ex.Message}",
                "XLDuck",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string GetEmbeddedRibbonXml()
    {
        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
        var resourceName = "XlDuck.Ribbon.xml";

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            throw new InvalidOperationException($"Embedded resource not found: {resourceName}");
        }

        using var reader = new System.IO.StreamReader(stream);
        return reader.ReadToEnd();
    }
}
