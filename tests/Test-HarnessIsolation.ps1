# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
$sentinel = $null
$sentinelSheet = $null
$sentinelRange = $null
$otherSheet = $null
$otherRange = $null
$otherBook = $null
$otherApplication = $null
try {
    $sentinel = New-XlDuckExcelSession
    $sentinelSheets = $sentinel.Workbook.Worksheets
    $sentinelSheet = $sentinelSheets.Item(1)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sentinelSheets)
    $sentinelRange = $sentinelSheet.Range('A1')
    $sentinelRange.Value2 = 'Unrelated workbook sentinel'
    $sentinelId = $sentinel.ProcessId
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-Tests.ps1') -NoBuild
    if ($LASTEXITCODE -ne 0) { throw 'Integration tests failed.' }
    Assert-XlDuckExcelSession $sentinel
    if ($sentinelRange.Value2 -ne 'Unrelated workbook sentinel') {
        throw 'The unrelated workbook was modified.'
    }
    Write-Host "PASS unrelated Excel PID $sentinelId and its workbook survived the suite."

    # A workbook subsequently added to a harness instance must survive cleanup too.
    $session = New-XlDuckExcelSession
    $otherBooks = $session.Application.Workbooks
    $otherBook = $otherBooks.Add()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBooks)
    $otherApplication = $otherBook.Application
    $otherSheets = $otherBook.Worksheets
    $otherSheet = $otherSheets.Item(1)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheets)
    $otherRange = $otherSheet.Range('A1')
    $otherRange.Value2 = 'Additional workbook sentinel'
    Close-XlDuckExcelSession $session
    if ($otherRange.Value2 -ne 'Additional workbook sentinel') {
        throw 'Cleanup touched the additional workbook.'
    }
    Write-Host 'PASS cleanup leaves additional workbooks open.'
} finally {
    if ($otherRange) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherRange); $otherRange = $null }
    if ($otherSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet); $otherSheet = $null }
    if ($otherBook) {
        $otherBook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook)
        $otherBook = $null
    }
    if ($otherApplication) {
        $otherApplication.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherApplication)
        $otherApplication = $null
    }
    if ($sentinelRange) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sentinelRange); $sentinelRange = $null }
    if ($sentinelSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sentinelSheet); $sentinelSheet = $null }
    if ($sentinel) { Close-XlDuckExcelSession $sentinel }
}
