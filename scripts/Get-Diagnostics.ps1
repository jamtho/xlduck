# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$package = [IO.Path]::GetFullPath($PackageDirectory)
$verification = 'failed'
try {
    $null = & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $package 6>&1
    $verification = 'passed'
} catch { } # Do not include exception messages, which may contain local paths.
$runtimes = @()
if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    $runtimes = @(dotnet --list-runtimes 2>$null | ForEach-Object {
        if ($_ -match '^(Microsoft\.[A-Za-z.]+) ([0-9A-Za-z.+-]+) \[') {
            [pscustomobject]@{name=$Matches[1];version=$Matches[2]}
        }
    })
}
$binaries = @('XlDuck.dll','XlDuck-AddIn64.xll','worker/XlDuck.Worker.exe') | ForEach-Object {
    $file = Join-Path $package $_
    [pscustomobject]@{
        name=$_
        present=[IO.File]::Exists($file)
        version=if ([IO.File]::Exists($file)) { [Diagnostics.FileVersionInfo]::GetVersionInfo($file).FileVersion } else { $null }
    }
}
[ordered]@{
    schemaVersion=1
    capturedUtc=[DateTime]::UtcNow.ToString('o')
    windowsVersion=[Environment]::OSVersion.Version.ToString()
    operatingSystem64Bit=[Environment]::Is64BitOperatingSystem
    diagnosticProcess64Bit=[Environment]::Is64BitProcess
    packageManifest=$verification
    binaries=@($binaries)
    availableDotnetRuntimes=$runtimes
    note='No workbook content, paths, settings, credentials or logs are collected. DuckStatus() reports the runtime actually loaded inside Excel.'
} | ConvertTo-Json -Depth 5
