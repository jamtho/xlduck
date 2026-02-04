// Copyright (c) 2026 James Thompson
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

namespace XlDuck;

/// <summary>
/// Simple file logger for debugging.
/// </summary>
public static class Log
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XlDuck",
        "xlduck.log");

    private static readonly object _lock = new();

    static Log()
    {
        try
        {
            var dir = Path.GetDirectoryName(LogPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
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
}
