// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck;

/// <summary>
/// File logger with per-session log files and weekly rotation.
/// Logs to %LOCALAPPDATA%\XlDuck\xlduck-{timestamp}.log
/// </summary>
public static class Log
{
    private static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XlDuck");

    private static readonly string LogPath = Path.Combine(
        LogDir,
        $"xlduck-{DateTime.Now:yyyyMMdd-HHmmss}.log");

    private static readonly object _lock = new();

    static Log()
    {
        try
        {
            if (!Directory.Exists(LogDir))
                Directory.CreateDirectory(LogDir);

            PurgeOldLogs();
        }
        catch { }
    }

    public static void Write(string message)
    {
        try
        {
            lock (_lock)
            {
                var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}";
                File.AppendAllText(LogPath, line + Environment.NewLine);
                System.Diagnostics.Debug.WriteLine(line);
            }
        }
        catch { }
    }

    public static void Error(string context, Exception ex)
    {
        Write($"[ERROR] {context}: {ex.Message}");
        Write($"        {ex.GetType().Name}: {ex.StackTrace}");
    }

    public static string GetLogPath() => LogPath;

    private static void PurgeOldLogs()
    {
        try
        {
            var cutoff = DateTime.Now.AddDays(-7);
            foreach (var file in Directory.GetFiles(LogDir, "xlduck-*.log"))
            {
                if (File.GetCreationTime(file) < cutoff)
                    File.Delete(file);
            }
        }
        catch { }
    }
}
