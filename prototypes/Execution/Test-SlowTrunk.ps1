# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([ValidateSet('inproc', 'worker')][string]$Mode = 'worker')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../../scripts/ExcelSession.ps1')
$session = $null
$sheet = $null
try {
    $session = New-XlDuckExcelSession -AddInPath (Join-Path $PSScriptRoot 'AddIn/bin/Debug/net8.0-windows/AddIn-AddIn64.xll')
    $excel = $session.Application
    $sheet = $session.Workbook.Worksheets.Item(1)
    $excel.Run('ProtoSetWorker', (Join-Path $PSScriptRoot 'Worker/bin/Debug/net8.0/Worker.exe')) | Out-Null
    $excel.Run('ProtoSetPublisher', 'targeted', $session.WorkbookName) | Out-Null
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $sheet.Range('A1').Formula2 = ('=ProtoQuery("{0}","SELECT 123 AS value FROM (SELECT sum(sin(i::DOUBLE)) AS s FROM range(3000000000) t(i)) WHERE s IS NOT NULL","")' -f $Mode)
    $sheet.Range('D1').Formula2 = ('=ProtoScalar("{0}",A1)' -f $Mode)
    $edit = [Diagnostics.Stopwatch]::StartNew()
    $sheet.Range('F1').Formula2 = '=6*7'
    [void](Wait-XlDuckCell $sheet 'F1' -TimeoutSeconds 3 -Until { param($v) $v -is [double] -and $v -eq 42 })
    $editMs = $edit.Elapsed.TotalMilliseconds
    $state = $excel.Run('ProtoSessionInfo', $session.WorkbookName, $Mode) | ConvertFrom-Json
    if ($state.Pending -eq 0) { throw 'The slow request was not pending during the ordinary Excel edit.' }
    Write-Host "Ordinary Excel edit completed in $([Math]::Round($editMs)) ms while the trunk was pending."
    [void](Wait-XlDuckCell $sheet 'D1' -TimeoutSeconds 120 -Until {
        param($v)
        if ($v -is [string] -and $v.StartsWith('#ERROR:')) { throw $v }
        $v -is [double] -and $v -eq 123
    })
    $result = [pscustomobject]@{ mode = $Mode; trunkMs = $watch.Elapsed.TotalMilliseconds; ordinaryExcelEditMs = $editMs; leaf = $sheet.Range('D1').Value2 }
    $result | ConvertTo-Json | Write-Output
    [void](New-Item -ItemType Directory -Path (Join-Path $PSScriptRoot '../../artifacts/execution') -Force)
    $result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot "../../artifacts/execution/$Mode-slow.json") -Encoding UTF8
} finally {
    if ($session -and $session.Application) {
        try { $session.Application.Run('ProtoShutdown') | Out-Null } catch { Write-Warning $_ }
    }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null }
    $excel = $null
    if ($session) { Close-XlDuckExcelSession $session }
}
