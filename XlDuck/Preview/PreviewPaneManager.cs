// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace XlDuck.Preview;

/// <summary>
/// Manages preview panes, one per Excel window.
/// </summary>
public class PreviewPaneManager : IDisposable
{
    private static PreviewPaneManager? _instance;
    private static readonly object _instanceLock = new();

    private readonly Dictionary<int, (CustomTaskPane Pane, PreviewController Controller)> _panes = new();
    private readonly object _panesLock = new();
    private dynamic? _excel;
    private SheetSelectionChangeHandler? _selectionChangeHandler;
    private bool _disposed;

    /// <summary>
    /// Get or create the singleton instance.
    /// </summary>
    public static PreviewPaneManager Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_instanceLock)
                {
                    _instance ??= new PreviewPaneManager();
                }
            }
            return _instance;
        }
    }

    private PreviewPaneManager()
    {
        Initialize();
    }

    private void Initialize()
    {
        try
        {
            Log.Write("[PaneManager] Initialize starting");
            _excel = ExcelDnaUtil.Application;
            Log.Write($"[PaneManager] Got Excel application: {_excel != null}");

            // Subscribe to selection change event using dynamic
            _selectionChangeHandler = new SheetSelectionChangeHandler(OnSheetSelectionChange);
            _excel.SheetSelectionChange += _selectionChangeHandler;

            Log.Write("[PaneManager] Initialized successfully");
        }
        catch (Exception ex)
        {
            Log.Error("PaneManager.Initialize", ex);
        }
    }

    // Delegate for SheetSelectionChange event
    private delegate void SheetSelectionChangeHandler(object sheet, dynamic target);

    /// <summary>
    /// Toggle the preview pane for the active window.
    /// </summary>
    public void TogglePane()
    {
        Log.Write("[PaneManager] TogglePane called");
        try
        {
            var hwnd = GetActiveWindowHwnd();
            Log.Write($"[PaneManager] Active window hwnd: {hwnd}");
            if (hwnd == 0)
            {
                Log.Write("[PaneManager] No active window, returning");
                return;
            }

            lock (_panesLock)
            {
                if (_panes.TryGetValue(hwnd, out var existing))
                {
                    Log.Write($"[PaneManager] Found existing pane, toggling visibility from {existing.Pane.Visible}");
                    // Toggle visibility
                    existing.Pane.Visible = !existing.Pane.Visible;

                    // If showing, refresh with current selection
                    if (existing.Pane.Visible)
                    {
                        RefreshCurrentSelection(existing.Controller);
                    }
                }
                else
                {
                    Log.Write("[PaneManager] Creating new pane");
                    // Create new pane
                    var pane = CreatePane(hwnd);
                    if (pane != null)
                    {
                        Log.Write("[PaneManager] Pane created, setting visible");
                        pane.Value.Pane.Visible = true;
                        RefreshCurrentSelection(pane.Value.Controller);
                    }
                    else
                    {
                        Log.Write("[PaneManager] Failed to create pane");
                    }
                }
            }
            Log.Write("[PaneManager] TogglePane completed");
        }
        catch (Exception ex)
        {
            Log.Error("PaneManager.TogglePane", ex);
        }
    }

    /// <summary>
    /// Check if the preview pane is visible for the active window.
    /// </summary>
    public bool IsPaneVisible()
    {
        try
        {
            var hwnd = GetActiveWindowHwnd();
            if (hwnd == 0) return false;

            lock (_panesLock)
            {
                if (_panes.TryGetValue(hwnd, out var existing))
                {
                    return existing.Pane.Visible;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewPaneManager] IsPaneVisible error: {ex.Message}");
        }
        return false;
    }

    private (CustomTaskPane Pane, PreviewController Controller)? CreatePane(int hwnd)
    {
        Log.Write($"[PaneManager] CreatePane for hwnd {hwnd}");
        try
        {
            Log.Write("[PaneManager] Calling CustomTaskPaneFactory.CreateCustomTaskPane with type");
            var taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(
                typeof(PreviewPane),
                "XlDuck Preview");

            Log.Write("[PaneManager] Setting dock position and width");
            taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = 875;

            Log.Write("[PaneManager] Getting content control");
            var previewPane = (PreviewPane)taskPane.ContentControl;

            Log.Write("[PaneManager] Creating PreviewController");
            var controller = new PreviewController(previewPane);

            var result = (taskPane, controller);
            _panes[hwnd] = result;

            Log.Write($"[PaneManager] Pane created successfully for hwnd {hwnd}");
            return result;
        }
        catch (Exception ex)
        {
            Log.Error("PaneManager.CreatePane", ex);
            return null;
        }
    }

    private void OnSheetSelectionChange(object sheet, dynamic target)
    {
        try
        {
            var hwnd = GetActiveWindowHwnd();
            if (hwnd == 0) return;

            (CustomTaskPane Pane, PreviewController Controller)? paneInfo;
            lock (_panesLock)
            {
                if (!_panes.TryGetValue(hwnd, out var info))
                    return;
                paneInfo = info;
            }

            // Only process if pane is visible
            if (!paneInfo.Value.Pane.Visible) return;

            // Get cell value for single cell selection
            string? cellValue = null;
            if (target.Count == 1)
            {
                var value = target.Value2;
                cellValue = value?.ToString();
            }

            paneInfo.Value.Controller.OnSelectionChanged(cellValue);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewPaneManager] OnSheetSelectionChange error: {ex.Message}");
        }
    }

    private void RefreshCurrentSelection(PreviewController controller)
    {
        try
        {
            dynamic? selection = _excel?.Selection;
            if (selection != null && (int)selection.Count == 1)
            {
                var value = selection.Value2;
                controller.RefreshNow(value?.ToString());
            }
            else
            {
                controller.RefreshNow(null);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewPaneManager] RefreshCurrentSelection error: {ex.Message}");
        }
    }

    private int GetActiveWindowHwnd()
    {
        try
        {
            if (_excel?.ActiveWindow != null)
            {
                return _excel.ActiveWindow.Hwnd;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewPaneManager] GetActiveWindowHwnd error: {ex.Message}");
        }
        return 0;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_excel != null && _selectionChangeHandler != null)
            {
                _excel.SheetSelectionChange -= _selectionChangeHandler;
            }

            lock (_panesLock)
            {
                foreach (var (_, (pane, controller)) in _panes)
                {
                    controller.Dispose();
                    try { pane.Delete(); } catch { }
                }
                _panes.Clear();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PreviewPaneManager] Dispose error: {ex.Message}");
        }
    }
}
