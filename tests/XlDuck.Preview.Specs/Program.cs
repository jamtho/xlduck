// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Runtime.InteropServices;
using System.Text.Json;
using XlDuck.Preview;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // This is a disposable standalone process, never the user's Excel.
        var context = args.Single() switch { "unaware" => -1, "system" => -2, "per-monitor" => -4, _ => throw new ArgumentException("Choose a DPI context.") };
        SetThreadDpiAwarenessContext(context);
        using var form = new Form { Text = "XlDuck owned DPI test: " + args[0], ClientSize = new(640, 480), AutoScaleMode = AutoScaleMode.None };
        using var browser = new PreviewBrowser { Dock = DockStyle.Fill };
        form.Controls.Add(browser);
        var exitCode = 1;
        using var timeout = new System.Windows.Forms.Timer { Interval = 30000 };
        timeout.Tick += (_, _) => { Console.Error.WriteLine("Preview DPI test timed out."); form.Close(); };
        form.Shown += async (_, _) =>
        {
            timeout.Start();
            try
            {
                var profile = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../artifacts/preview-dpi-profile"));
                await browser.InitializeAsync(profile);
                var core = browser.CoreWebView2!;
                var loaded = new TaskCompletionSource();
                core.NavigationCompleted += (_, e) => { if (e.IsSuccess) loaded.TrySetResult(); else loaded.TrySetException(new Exception("Test page failed.")); };
                core.NavigateToString("<html><body style='font:14px Segoe UI'><h2>XlDuck DPI test</h2>One CSS pixel follows the host window.</body></html>");
                await loaded.Task;
                async Task Check()
                {
                    var watch = System.Diagnostics.Stopwatch.StartNew();
                    while (true)
                    {
                        using var metrics = JsonDocument.Parse(await core.ExecuteScriptAsync("({dpr:devicePixelRatio,width:innerWidth})"));
                        var expected = GetDpiForWindow(browser.Handle) / 96d;
                        var dpr = metrics.RootElement.GetProperty("dpr").GetDouble();
                        var width = metrics.RootElement.GetProperty("width").GetDouble();
                        if (Math.Abs(dpr - expected) <= .01 && Math.Abs(width * expected - browser.ClientSize.Width) <= 3 && browser.ZoomFactor == 1)
                        {
                            Console.WriteLine($"PASS {args[0]}: pixel ratio {dpr}, CSS width {width}, client {browser.ClientSize.Width}.");
                            return;
                        }
                        if (watch.Elapsed.TotalSeconds > 3)
                            throw new Exception($"Double scaling: expected {expected}, actual {dpr}, CSS width {width}, client {browser.ClientSize.Width}.");
                        await Task.Delay(50);
                    }
                }
                await Check();
                form.ClientSize = new(480, 360);
                await Check();
                exitCode = 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error); }
            finally { timeout.Stop(); form.Close(); }
        };
        Application.Run(form);
        return exitCode;
    }
    [DllImport("user32.dll")] private static extern nint SetThreadDpiAwarenessContext(nint context);
    [DllImport("user32.dll")] private static extern uint GetDpiForWindow(nint window);
}

namespace XlDuck
{
    internal static class Log
    {
        internal static void Write(string message) => Console.WriteLine(message);
        internal static void Error(string message, Exception error) => Console.Error.WriteLine(message + ": " + error);
    }
}
