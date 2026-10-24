// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0

// Only the UI host and logging are replaced. Tests execute the production computational code.
namespace XlDuck
{
    public static class Log
    {
        public static void Write(string _) { }
        public static void Error(string _, Exception __) { }
    }
    public sealed class RibbonController
    {
        public static RibbonController? Instance => null;
        public void InvalidatePreviewToggle() { }
    }
}
namespace XlDuck.Preview
{
    public sealed class PreviewPaneManager
    {
        public static PreviewPaneManager Instance { get; } = new();
        public void TogglePane() { }
    }
}
