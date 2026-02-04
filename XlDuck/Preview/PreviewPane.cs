// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace XlDuck.Preview;

/// <summary>
/// Empty interface for COM default interface (required for .NET 6+ CTP support).
/// </summary>
[ComVisible(true)]
[Guid("F9A5E8C1-2B3D-4E6F-A7B8-C9D0E1F2A3B4")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPreviewPane { }

/// <summary>
/// UserControl hosting WebView2 for preview display.
/// </summary>
[ComVisible(true)]
[Guid("E8F4D7B2-3A1C-4E9F-8B5D-2C7A6F0E1D3B")]
[ProgId("XlDuck.PreviewPane")]
[ComDefaultInterface(typeof(IPreviewPane))]
[ClassInterface(ClassInterfaceType.None)]
public class PreviewPane : UserControl, IPreviewPane
{
    private WebView2? _webView;
    private Label? _fallbackLabel;
    private bool _isWebViewReady;
    private string? _pendingJson;

    public PreviewPane()
    {
        Log.Write("[PreviewPane] Constructor called");
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Log.Write("[PreviewPane] InitializeComponent");
        SuspendLayout();

        BackColor = System.Drawing.Color.White;
        Dock = DockStyle.Fill;

        ResumeLayout(false);
        Log.Write("[PreviewPane] InitializeComponent done");
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        Log.Write("[PreviewPane] OnLoad");

        if (DesignMode) return;

        await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        Log.Write("[PreviewPane] InitializeWebViewAsync starting");
        try
        {
            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_webView);

            // Use a custom user data folder in %LOCALAPPDATA% to avoid permission issues
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XlDuck", "WebView2");
            Directory.CreateDirectory(userDataFolder);

            Log.Write($"[PreviewPane] Using WebView2 user data folder: {userDataFolder}");
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

            Log.Write("[PreviewPane] Calling EnsureCoreWebView2Async");
            await _webView.EnsureCoreWebView2Async(env);
            Log.Write("[PreviewPane] WebView2 core initialized");

            // Load embedded HTML
            var html = LoadEmbeddedHtml();
            _webView.NavigateToString(html);

            _webView.NavigationCompleted += OnNavigationCompleted;
        }
        catch (WebView2RuntimeNotFoundException)
        {
            Log.Write("[PreviewPane] WebView2 runtime not found");
            ShowFallbackMessage(
                "WebView2 Runtime not installed.\n\n" +
                "Download from:\nhttps://developer.microsoft.com/microsoft-edge/webview2/");
        }
        catch (Exception ex)
        {
            Log.Error("PreviewPane.InitializeWebViewAsync", ex);
            ShowFallbackMessage($"Failed to initialize preview:\n{ex.Message}");
        }
    }

    private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        Log.Write($"[PreviewPane] NavigationCompleted, success={e.IsSuccess}");
        _isWebViewReady = true;

        // Send any pending state
        if (_pendingJson != null)
        {
            PostMessage(_pendingJson);
            _pendingJson = null;
        }
    }

    /// <summary>
    /// Set the preview state (sends JSON to WebView2).
    /// </summary>
    public void SetState(PreviewModel model)
    {
        Log.Write($"[PreviewPane] SetState: {model.Kind}");

        var json = model.ToJson();

        if (_isWebViewReady && _webView?.CoreWebView2 != null)
        {
            PostMessage(json);
        }
        else
        {
            // Queue for when WebView2 is ready
            _pendingJson = json;
        }
    }

    private void PostMessage(string json)
    {
        try
        {
            _webView?.CoreWebView2?.PostWebMessageAsString(json);
        }
        catch (Exception ex)
        {
            Log.Error("PreviewPane.PostMessage", ex);
        }
    }

    private void ShowFallbackMessage(string message)
    {
        if (_fallbackLabel != null) return;

        _webView?.Dispose();
        _webView = null;
        Controls.Clear();

        _fallbackLabel = new Label
        {
            Text = message,
            Dock = DockStyle.Fill,
            TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
            ForeColor = System.Drawing.Color.FromArgb(180, 0, 0),
            BackColor = System.Drawing.Color.FromArgb(255, 248, 248),
            Padding = new Padding(20),
            Font = new System.Drawing.Font("Segoe UI", 10)
        };
        Controls.Add(_fallbackLabel);
    }

    private static string LoadEmbeddedHtml()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = "XlDuck.Preview.preview.html";

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            throw new InvalidOperationException($"Embedded resource not found: {resourceName}");
        }

        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _webView?.Dispose();
        }
        base.Dispose(disposing);
    }
}
