# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0

param([switch]$NoBuild)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelSession.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $NoBuild) { Build-XlDuckAddIn }
$addInPath = Join-Path $repoRoot 'XlDuck/bin/Debug/net8.0-windows/XlDuck-AddIn64.xll'
# A private copy keeps subsequent builds independent of DLLs loaded by this instance.
$developmentCopy = Copy-XlDuckDevelopmentBuild $addInPath
$session = New-XlDuckExcelSession -AddInPath $developmentCopy -Visible
Write-Host "Opened development Excel PID $($session.ProcessId), workbook $($session.WorkbookName)." -ForegroundColor Green
Write-Host 'Existing Excel instances were left alone. Close this development instance yourself when finished.'
Write-Host "Development build: $developmentCopy"
# Leave the visible development instance under the user's control.
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Workbook)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Application)
