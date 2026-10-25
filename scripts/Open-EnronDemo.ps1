# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([string]$WorkbookPath, [string]$AddInPath, [string]$ConnectionFile)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DemoSession.ps1')
if (-not $WorkbookPath) {
    $file = Get-ChildItem -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs') -Filter 'XlDuck Enron Demo.xlsx' -Recurse -File |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    if (-not $file) { throw 'Create the Enron demo or supply -WorkbookPath.' }
    $WorkbookPath = $file.FullName
}
$session = $null; $sheet = $null; $handedOver = $false
try {
    # Open the new demo at its configured path so its saved connections apply.
    # The legacy source workbook is never opened or changed by this launcher.
    $session = New-XlDuckDemoSession -WorkbookPath $WorkbookPath -AddInPath $AddInPath
    if (Invoke-XlDuckExcel { $session.Workbook.ReadOnly }) { throw 'The demo is already open or read-only. Use its existing Excel window.' }
    $sheet = Invoke-XlDuckExcel { $session.Workbook.Worksheets.Item('Research') }
    Invoke-XlDuckExcel { $sheet.Range('J1').Formula2 = '=DuckStatus()' }
    [void](Wait-XlDuckCell $sheet 'J1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).ready })
    if ($ConnectionFile) {
        $settings = Get-Content -LiteralPath $ConnectionFile -Raw
        Invoke-XlDuckExcel { $session.Application.Run('DuckSetConnections',$settings,$session.Workbook.Name) | Out-Null }
    }
    Invoke-XlDuckExcel { $sheet.Range('J1').Formula2 = '=DuckInspect(C12,"rows")' }
    [void](Wait-XlDuckCell $sheet 'J1' -Until { param($v) $v -is [double] -and $v -ge 0 } -TimeoutSeconds 60)
    Invoke-XlDuckExcel {
        $sheet.Range('J1').ClearContents()
        $sheet.Activate(); $sheet.Range('C12').Select() | Out-Null
        $session.Application.ActiveWindow.ScrollColumn = 1
        $session.Application.ActiveWindow.ScrollRow = 1
        $session.Workbook.Save()
    }
    Invoke-XlDuckExcel { $session.Application.Visible = $true }
    $handedOver = $true
    Invoke-XlDuckExcel { $session.Application.WindowState = -4137; $session.Application.Run('DuckTogglePreview') | Out-Null }
    Write-Host "Enron demo opened in its own Excel process (PID $($session.ProcessId))."
    Write-Host "Workbook: $($session.Workbook.FullName)"
    Write-Host 'Edit Research!C6:C8. Select C12/C13 for charts, C14 for records, or edit SQL on Queries.'
} finally {
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session -and -not $handedOver) { Close-XlDuckExcelSession $session }
    elseif ($session) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Workbook)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Application)
        $session.Process.Dispose()
    }
}
