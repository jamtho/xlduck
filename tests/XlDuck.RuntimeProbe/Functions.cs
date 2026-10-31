// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System;
using ExcelDna.Integration;

public static class RuntimeProbe
{
#if PROBE_Framework
    [ExcelFunction(Name = "XlDuckProbeFramework")]
    public static string Version() => Environment.Version.ToString();
    // Test-only bootstrap: this changes the owned Excel process, never user/machine settings.
    // .NET Framework can run before CoreCLR without choosing the latter's runtime.
    [ExcelCommand(Name = "XlDuckProbeRuntimeRoot")]
    public static void RuntimeRoot(string directory)
    {
        if (!System.IO.File.Exists(System.IO.Path.Combine(directory, "dotnet.exe"))) throw new ArgumentException("Missing portable runtime.");
        Environment.SetEnvironmentVariable("DOTNET_ROOT", directory);
        Environment.SetEnvironmentVariable("DOTNET_ROOT_X64", directory);
    }
#elif PROBE_Eight
    [ExcelFunction(Name = "XlDuckProbeEight")]
    public static string Version() => Environment.Version.ToString();
#else
    [ExcelFunction(Name = "XlDuckProbeTen")]
    public static string Version() => Environment.Version.ToString();
#endif
}
