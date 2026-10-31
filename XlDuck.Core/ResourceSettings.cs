// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Core;

/// <summary>Per-worker budgets, fixed when the host starts the workbook session.</summary>
public sealed record ResourceSettings(int MemoryMiB = 1024, int SpillGiB = 4, int Threads = 2)
{
    public void Validate()
    {
        if (MemoryMiB is < 128 or > 65536) throw new ArgumentException("Memory must be between 128 and 65536 MiB.");
        if (SpillGiB is < 1 or > 1024) throw new ArgumentException("Spill disk must be between 1 and 1024 GiB.");
        if (Threads is < 1 or > 16) throw new ArgumentException("Threads must be between 1 and 16.");
    }
}
