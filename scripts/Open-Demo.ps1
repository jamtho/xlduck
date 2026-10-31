# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([string]$WorkbookPath, [string]$AddInPath, [switch]$UseWorkbookInPlace)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DemoSession.ps1')
if (-not $WorkbookPath) {
    $WorkbookPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'examples/research-demo/XlDuck Research Demo.xlsx'
    if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw 'Create the research demo or supply -WorkbookPath.' }
}
$session = $null; $sheet = $null; $handedOver = $false
try {
    $session = New-XlDuckDemoSession -WorkbookPath $WorkbookPath -AddInPath $AddInPath -CopyWorkbook:(-not $UseWorkbookInPlace)
    $sheet = $session.Workbook.Worksheets.Item('Research')
    $sheet.Range('C47').Formula2 = '=DuckStatus()'
    $status = Wait-XlDuckCell $sheet 'C47' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).ready }
    # Resolve through DuckInspect so a saved token cannot be mistaken for a live result.
    $sheet.Range('C47').Formula2 = '=DuckInspect(C11,"rows")'
    [void](Wait-XlDuckCell $sheet 'C47' -Until { param($v) $v -is [double] -and $v -gt 0 })
    $sheet.Range('C47').ClearContents()
    Invoke-XlDuckExcel { $session.Workbook.Save() }
    Invoke-XlDuckExcel { $sheet.Activate(); $sheet.Range('C11').Select() | Out-Null }
    $session.Application.Visible = $true
    $handedOver = $true
    $session.Application.WindowState = -4137
    $session.Application.Run('DuckTogglePreview') | Out-Null
    Write-Host "Demo opened in its own Excel process (PID $($session.ProcessId))."
    Write-Host "Working copy: $($session.Workbook.FullName)"
    Write-Host 'Edit Research!C7:C9, or the observations. Close this Excel instance normally when finished.'
} finally {
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session -and -not $handedOver) { Close-XlDuckExcelSession $session }
    elseif ($session) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Workbook)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Application)
        $session.Process.Dispose()
    }
}
