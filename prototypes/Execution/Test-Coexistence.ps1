# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../../scripts/ExcelSession.ps1')
$legacy = (Resolve-Path (Join-Path $PSScriptRoot '../../XlDuck/bin/Debug/net8.0-windows/XlDuck-AddIn64.xll')).Path
$prototype = (Resolve-Path (Join-Path $PSScriptRoot 'AddIn/bin/Debug/net8.0-windows/AddIn-AddIn64.xll')).Path
foreach ($order in @('legacy-first', 'prototype-first')) {
    $session = $null
    $sheet = $null
    try {
        $session = New-XlDuckExcelSession
        $paths = if ($order -eq 'legacy-first') { @($legacy, $prototype) } else { @($prototype, $legacy) }
        foreach ($path in $paths) {
            if (-not $session.Application.RegisterXLL($path)) { throw "Could not load $path" }
        }
        $session.Application.Run('ProtoSetWorker', (Join-Path $PSScriptRoot 'Worker/bin/Debug/net8.0/Worker.exe')) | Out-Null
        $session.Application.Run('ProtoSetPublisher', 'targeted', $session.WorkbookName) | Out-Null
        $sheet = $session.Workbook.Worksheets.Item(1)
        $sheet.Range('A1').Formula2 = '=DuckQueryOutScalar("SELECT 42")'
        $sheet.Range('B1').Formula2 = '=ProtoQuery("worker","SELECT 84 AS value","")'
        $sheet.Range('C1').Formula2 = '=ProtoScalar("worker",B1)'
        [void](Wait-XlDuckCell $sheet 'A1' -Until { param($v) $v -is [double] -and $v -eq 42 })
        [void](Wait-XlDuckCell $sheet 'C1' -Until { param($v) $v -is [double] -and $v -eq 84 })
        Write-Host "PASS ${order}: legacy and prototype calculations coexist on installed .NET 8."
    } finally {
        if ($session -and $session.Application) {
            try { $session.Application.Run('ProtoShutdown') | Out-Null } catch { Write-Warning $_ }
        }
        if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null }
        if ($session) { Close-XlDuckExcelSession $session }
    }
}
