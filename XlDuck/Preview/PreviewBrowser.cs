// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Runtime.InteropServices;
using Microsoft.Web.WebView2.Core;

namespace XlDuck.Preview;

/// <summary>Windowed WebView with explicit Office child-window DPI ownership.</summary>
internal sealed class PreviewBrowser : Control
{
    private CoreWebView2Controller? _controller;
    private uint _dpi;
    internal CoreWebView2? CoreWebView2 => _controller?.CoreWebView2;
    internal double ZoomFactor => _controller?.ZoomFactor ?? 1;

    internal PreviewBrowser() => TabStop = true;

    internal async Task<CoreWebView2Environment> InitializeAsync(string root)
    {
        var window = Handle;
        var context = GetWindowDpiAwarenessContext(window);
        // WebView browser processes cannot share a profile across incompatible
        // DPI awareness modes. Keep their persistent caches separate.
        var folder = Path.Combine(root, "dpi-" + GetAwarenessFromDpiAwarenessContext(context));
        Task<CoreWebView2Environment> environmentTask;
        using (new DpiContext(context)) environmentTask = CoreWebView2Environment.CreateAsync(null, folder);
        var environment = await environmentTask;
        if (IsDisposed || Disposing) return environment;
        Task<CoreWebView2Controller> controllerTask;
        // Restore the thread context before awaiting: Excel owns the UI thread.
        using (new DpiContext(context)) controllerTask = environment.CreateCoreWebView2ControllerAsync(window);
        var controller = await controllerTask;
        if (IsDisposed || Disposing) { controller.Close(); return environment; }
        _controller = controller;
        controller.ShouldDetectMonitorScaleChanges = false;
        controller.BoundsMode = CoreWebView2BoundsMode.UseRawPixels;
        controller.ZoomFactor = 1;
        controller.CoreWebView2.Settings.IsZoomControlEnabled = false;
        controller.CoreWebView2.Settings.IsPinchZoomEnabled = false;
        controller.MoveFocusRequested += (_, args) =>
        {
            args.Handled = true;
            Parent?.SelectNextControl(this, args.Reason == CoreWebView2MoveFocusReason.Next, true, true, true);
        };
        UpdateLayout();
        return environment;
    }

    private void UpdateLayout()
    {
        if (_controller is null || IsDisposed || Disposing || !IsHandleCreated) return;
        try { UpdateControllerLayout(); }
        catch (Exception error) when (error is COMException or InvalidOperationException)
        { Log.Error("Preview window layout", error); }
    }

    private void UpdateControllerLayout()
    {
        // A system-aware/unaware Office child can be bitmap-scaled by Windows.
        // Applying the monitor scale again inside Chromium doubles that scaling.
        var dpi = GetDpiForWindow(Handle);
        if (dpi == 0) return;
        _controller!.RasterizationScale = dpi / 96d;
        _controller.Bounds = ClientRectangle;
        _controller.IsVisible = Visible;
        _controller.NotifyParentWindowPositionChanged();
        if (_dpi != dpi)
        {
            _dpi = dpi;
            Log.Write($"[PreviewBrowser] WindowDPI={dpi}, RasterizationScale={_controller.RasterizationScale:F3}");
        }
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); UpdateLayout(); }
    protected override void OnVisibleChanged(EventArgs e) { base.OnVisibleChanged(e); UpdateLayout(); }
    protected override void OnLocationChanged(EventArgs e) { base.OnLocationChanged(e); UpdateLayout(); }
    protected override void OnEnter(EventArgs e)
    {
        base.OnEnter(e);
        try { _controller?.MoveFocus(CoreWebView2MoveFocusReason.Programmatic); }
        catch (Exception error) when (error is COMException or InvalidOperationException) { Log.Error("Preview focus", error); }
    }
    protected override void WndProc(ref Message message)
    {
        base.WndProc(ref message);
        if (message.Msg is 0x02E0 or 0x02E3 or 0x0047) UpdateLayout(); // DPI/parent-DPI/window position changed.
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            try { _controller?.Close(); }
            catch (Exception error) when (error is COMException or InvalidOperationException) { Log.Error("Preview browser cleanup", error); }
            finally { _controller = null; }
        }
        base.Dispose(disposing);
    }

    private sealed class DpiContext(nint context) : IDisposable
    {
        private readonly nint _previous = SetThreadDpiAwarenessContext(context);
        public void Dispose() { if (_previous != 0) SetThreadDpiAwarenessContext(_previous); }
    }
    [DllImport("user32.dll")] private static extern uint GetDpiForWindow(nint window);
    [DllImport("user32.dll")] private static extern nint GetWindowDpiAwarenessContext(nint window);
    [DllImport("user32.dll")] private static extern int GetAwarenessFromDpiAwarenessContext(nint context);
    [DllImport("user32.dll")] private static extern nint SetThreadDpiAwarenessContext(nint context);
}
