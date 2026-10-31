# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Production integration entry point. Every test owns its Excel process and workbook.
param([switch]$NoBuild, [string]$AddInPath, [switch]$WithPreview, [string]$RuntimeRoot, [string]$FrameworkProbe)
$ErrorActionPreference = 'Stop'
$options = @()
if ($NoBuild) { $options += '-NoBuild' }
if ($AddInPath) { $options += @('-AddInPath', $AddInPath) }
if ($RuntimeRoot) { $options += @('-RuntimeRoot', $RuntimeRoot, '-FrameworkProbe', $FrameworkProbe) }
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Production.ps1') @options
if ($LASTEXITCODE -ne 0) { throw 'Production integration failed.' }
$protectedOptions = @()
if ($AddInPath) { $protectedOptions += @('-AddInPath', $AddInPath) }
if ($RuntimeRoot) { $protectedOptions += @('-RuntimeRoot', $RuntimeRoot, '-FrameworkProbe', $FrameworkProbe) }
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-ProtectedView.ps1') @protectedOptions
if ($LASTEXITCODE -ne 0) { throw 'Protected View integration failed.' }
if ($WithPreview) {
    & dotnet build (Join-Path $PSScriptRoot 'XlDuck.Preview.Specs') --nologo -m:1 /nodeReuse:false /p:UseSharedCompilation=false /p:RestoreLockedMode=true
    if ($LASTEXITCODE -ne 0) { throw 'Preview DPI test build failed.' }
    foreach ($mode in @('unaware', 'system', 'per-monitor')) {
        & dotnet run --project (Join-Path $PSScriptRoot 'XlDuck.Preview.Specs') --no-build --no-restore -- $mode
        if ($LASTEXITCODE -ne 0) { throw "Preview DPI verification failed: $mode" }
    }
    $previewOptions = @('-NoBuild')
    if ($AddInPath) { $previewOptions += @('-AddInPath', $AddInPath) }
    if ($RuntimeRoot) { $previewOptions += @('-RuntimeRoot', $RuntimeRoot, '-FrameworkProbe', $FrameworkProbe) }
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-PreviewHost.ps1') @previewOptions
    if ($LASTEXITCODE -ne 0) { throw 'Native preview integration failed.' }
}
