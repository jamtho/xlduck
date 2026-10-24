# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param(
    [ValidateSet('inproc', 'worker')][string]$Mode = 'inproc',
    [ValidateSet('rtd', 'targeted')][string]$Publisher = 'rtd',
    [int]$Depth = 8,
    [int]$Repeats = 3,
    [switch]$NoBuild
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../../scripts/ExcelSession.ps1')
if (-not $NoBuild) {
    Build-XlDuckAddIn (Join-Path $PSScriptRoot 'Worker/Worker.csproj')
    Build-XlDuckAddIn (Join-Path $PSScriptRoot 'AddIn/AddIn.csproj')
}
$session = $null
$sheet = $null
$otherBook = $null
$results = [Collections.Generic.List[object]]::new()
function Set-CellFormula([string]$address, [string]$formula) {
    Invoke-XlDuckExcel { $sheet.Range($address).Formula2 = $formula }
}
function Wait-Number([string]$address, [double]$expected, [int]$timeout = 90) {
    [void](Wait-XlDuckCell $sheet $address -TimeoutSeconds $timeout -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -is [double] -and $value -eq $expected
    })
}
try {
    Write-Host "Execution experiment: $Mode, $Publisher, depth $Depth, $Repeats edits"
    $session = New-XlDuckExcelSession -AddInPath (Join-Path $PSScriptRoot 'AddIn/bin/Debug/net8.0-windows/AddIn-AddIn64.xll')
    $excel = $session.Application
    $sheet = $session.Workbook.Worksheets.Item(1)
    $excel.Run('ProtoSetWorker', (Join-Path $PSScriptRoot 'Worker/bin/Debug/net8.0/Worker.exe')) | Out-Null
    $excel.Run('ProtoSetPublisher', $Publisher, $session.WorkbookName) | Out-Null
    $throttle = $excel.RTD.ThrottleInterval
    Write-Host "RTD throttle unchanged at $throttle ms"
    # Never modify ThrottleInterval or calculation mode for the automatic-mode comparison.
    if ($excel.Calculation -ne -4105) { throw 'Automatic calculation is required for the main experiment.' }
    $sheet.Range('B1').Value2 = 1
    $watch = [Diagnostics.Stopwatch]::StartNew()
    Set-CellFormula 'A1' ('=ProtoQuery("{0}","SELECT $value AS value","",B1)' -f $Mode)
    for ($i = 2; $i -le $Depth + 1; $i++) {
        Set-CellFormula "A$i" ('=ProtoQuery("{0}","SELECT value+1 AS value FROM __input",A{1})' -f $Mode, ($i-1))
    }
    Set-CellFormula 'D1' ('=ProtoScalar("{0}",A{1})' -f $Mode, ($Depth+1))
    Wait-Number 'D1' ($Depth+1)
    $coldMs = $watch.Elapsed.TotalMilliseconds
    Write-Host "Cold chain: $([Math]::Round($coldMs)) ms"
    $samples = @()
    for ($run = 0; $run -lt $Repeats; $run++) {
        $value = $run + 10
        $watch.Restart()
        Invoke-XlDuckExcel { $sheet.Range('B1').Value2 = $value }
        Wait-Number 'D1' ($value+$Depth)
        $samples += $watch.Elapsed.TotalMilliseconds
        Write-Host "Edit $run completed in $([Math]::Round($watch.Elapsed.TotalMilliseconds)) ms"
    }
    Set-CellFormula 'E1' ('=ProtoExecutions("{0}",D1)' -f $Mode)
    $before = $sheet.Range('E1').Value2
    $handleBefore = $sheet.Range('A1').Value2
    $watch.Restart()
    $excel.CalculateFull()
    Wait-Number 'D1' (($Repeats-1)+10+$Depth)
    $cacheMs = $watch.Elapsed.TotalMilliseconds
    $sheet.Range('E1').Calculate()
    if ($sheet.Range('E1').Value2 -ne $before) { throw 'Full calculation reran unchanged queries.' }
    # F2/Enter is represented by assigning the exact same Formula2.
    $unchanged = $sheet.Range('A1').Formula2
    Set-CellFormula 'A1' $unchanged
    if ($sheet.Range('A1').Value2 -ne $handleBefore) { throw 'Unchanged formula reentry lost its handle.' }

    Set-CellFormula 'G1' ('=ProtoOut("{0}",A{1})' -f $Mode, ($Depth+1))
    Wait-Number 'G2' (($Repeats-1)+10+$Depth)
    Set-CellFormula 'I1' ('=ProtoPreview("{0}",A{1})' -f $Mode, ($Depth+1))
    [void](Wait-XlDuckCell $sheet 'I1' -TimeoutSeconds 30 -Until { param($v) $v -is [string] -and $v.StartsWith('{') })

    # A second workbook with identical formulas must have a different session and handle.
    $otherBook = $excel.Workbooks.Add()
    $otherSheet = $otherBook.Worksheets.Item(1)
    $otherSheet.Range('A1').Formula2 = ('=ProtoQuery("{0}","SELECT $value AS value","",{1})' -f $Mode, (($Repeats-1)+10))
    $otherHandle = Wait-XlDuckCell $otherSheet 'A1' -TimeoutSeconds 30 -Until { param($v) $v -is [string] -and $v.StartsWith('proto://') }
    if ($otherHandle -eq $handleBefore) { throw 'Identical requests in different workbooks share a session.' }
    $session.Workbook.Activate()
    $excel.Run('ProtoRestart', $session.WorkbookName) | Out-Null
    Write-Host "Restart invoked for $($session.WorkbookName). Diagnostics: $($excel.Run('ProtoDiagnostics'))"
    [void](Wait-XlDuckCell $sheet 'A1' -TimeoutSeconds 5 -Until { param($v) $v -is [string] -and $v.StartsWith('proto://') -and $v -ne $handleBefore })
    Wait-Number 'D1' (($Repeats-1)+10+$Depth)
    if ($sheet.Range('A1').Value2 -eq $handleBefore) { throw 'Restart did not replace the session.' }
    $otherSheet.Range('A1').Calculate()
    if ($otherSheet.Range('A1').Value2 -ne $otherHandle) { throw 'Restart changed another workbook.' }
    $otherBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook)
    $otherBook = $null

    $results.Add([pscustomobject]@{
        mode = $Mode; publisher = $Publisher; depth = $Depth; repeats = $Repeats
        rtdThrottleMs = $throttle; coldMs = $coldMs; editSamplesMs = $samples; unchangedRecalcMs = $cacheMs
        checks = @('chain', 'changed input', 'full recalc reuse', 'formula reentry reuse', 'scalar', 'spill', 'preview data', 'workbook isolation', 'restart')
    })
    $results | ConvertTo-Json -Depth 6 | Write-Output
    $artifactDir = Join-Path $PSScriptRoot '../../artifacts/execution'
    [void](New-Item -ItemType Directory -Path $artifactDir -Force)
    $results | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $artifactDir "$Mode-$Publisher.json") -Encoding UTF8
} finally {
    if ($session -and $session.Application) {
        try { Write-Host "Prototype diagnostics: $($session.Application.Run('ProtoDiagnostics'))" } catch { Write-Warning $_ }
        try { $session.Application.Run('ProtoShutdown') | Out-Null } catch { Write-Warning $_ }
    }
    if ($otherBook) { $otherBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null }
    $excel = $null
    if ($session) { Close-XlDuckExcelSession $session }
}
