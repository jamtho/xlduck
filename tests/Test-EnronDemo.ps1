# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Opt-in live test: requires the user's Enron S3 connection and saved credential.
param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][string]$ConnectionFile,
    [string]$AddInPath,
    [switch]$Prepare,
    [switch]$ReopenOnly
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/DemoSession.ps1')
$repo = Split-Path $PSScriptRoot -Parent
$support = Join-Path $repo 'artifacts/enron-demo'
$contract = Get-Content (Join-Path $support 'contract.json') -Raw | ConvertFrom-Json
$connection = Get-Content -LiteralPath $ConnectionFile -Raw
$session = $null; $research = $null; $queries = $null; $check = $null
function Wait-Number([string]$Cell, [double]$Expected) {
    [void](Wait-XlDuckCell $research $Cell -Until { param($v) $v -is [double] -and [Math]::Abs($v - $Expected) -lt 1e-10 } -TimeoutSeconds 60)
}
function Read-Stats {
    Invoke-XlDuckExcel { $check.Range('K1').Formula2 = '=DuckStatus()'; $check.Range('K1').Calculate() }
    $check.Range('K1').Value2 | ConvertFrom-Json
}
function Verify-Selection {
    # An independent scan of the source verifies the displayed aggregate and filters.
    $check.Range('A1').Value2 = [string]$research.Range('C5').Value2
    $check.Range('A2').Value2 = [double]$research.Range('C6').Value2
    $check.Range('A3').Value2 = [string]$research.Range('C7').Value2
    $check.Range('A4').Value2 = [string]$research.Range('C8').Value2
    $check.Range('A5').Value2 = @'
SELECT count(*) AS records,
       count(DISTINCT lower(coalesce(nullif(trim(from_email), ''), nullif(trim(from_name), ''), '(unknown)'))) AS senders,
       avg(coalesce(is_reply, false)::INTEGER) AS reply_share
FROM read_parquet(?)
WHERE year(timezone('UTC', date)) = ?
  AND (? = 'All' OR contains(lower(coalesce(subject_clean, subject, '')), lower(?)))
  AND (? = 'All' OR lower(coalesce(nullif(trim(from_email), ''), nullif(trim(from_name), ''), '(unknown)')) = lower(?))
'@
    $check.Range('D1').Formula2 = '=DuckQueryOut(A5,A1,A2,A3,A3,A4,A4)'
    [void](Wait-XlDuckCell $check 'D2' -Until { param($v) $v -is [double] } -TimeoutSeconds 60)
    $expected = [double]$check.Range('D2').Value2
    Wait-Number 'F7' $expected
    Wait-Number 'F8' ([double]$check.Range('E2').Value2)
    if ($expected -gt 0) { Wait-Number 'F9' ([double]$check.Range('F2').Value2) }
    else { [void](Wait-XlDuckCell $research 'F9' -Until { param($v) $null -eq $v -or $v -ceq '' }) }
    # Reconcile all monthly rows to the same count, and verify both live plot handles.
    $check.Range('K2').Formula2 = '=SUM(Research!C19:C30)'
    [void](Wait-XlDuckCell $check 'K2' -Until { param($v) $v -is [double] -and $v -eq $expected })
    foreach ($cell in @('C12','C13')) {
        $check.Range('K3').Formula2 = '=DuckInspect(Research!' + $cell + ',"rows")'
        [void](Wait-XlDuckCell $check 'K3' -Until { param($v) $v -is [double] -and ($expected -eq 0 -or $v -gt 0) })
    }
    return $expected
}
try {
    $session = New-XlDuckDemoSession -WorkbookPath $WorkbookPath -AddInPath $AddInPath -CopyWorkbook:(-not ($Prepare -or $ReopenOnly))
    $research = Invoke-XlDuckExcel { $session.Workbook.Worksheets.Item('Research') }
    $queries = Invoke-XlDuckExcel { $session.Workbook.Worksheets.Item('Queries') }
    $check = Invoke-XlDuckExcel { $session.Workbook.Worksheets.Add() }
    Invoke-XlDuckExcel { $check.Name = 'Verification'; $check.Range('K1').Formula2 = '=DuckStatus()' }
    [void](Wait-XlDuckCell $check 'K1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).ready })
    if (-not $ReopenOnly) { $session.Application.Run('DuckSetConnections',$connection,$session.Workbook.Name) | Out-Null }
    if ($Prepare) {
        foreach ($tab in $contract.formulas.PSObject.Properties) {
            $target = if ($tab.Name -eq 'Queries') { $queries } else { $research }
            foreach ($formula in $tab.Value.PSObject.Properties) { Invoke-XlDuckExcel { $target.Range($formula.Name).Formula2 = $formula.Value } }
        }
    }
    $baseline = Verify-Selection
    if ($baseline -le 0) { throw 'Expected records for the default research selection.' }
    $root = $queries.Range('C6').Value2
    [void](Wait-XlDuckCell $research 'F11' -Until { param($v) $v -is [double] -and $v -gt 0 })
    Write-Host "PASS live Enron default: $($research.Range('F11').Value2) loaded records; $baseline matching records; monthly totals, sender count, reply share and both plots verified."
    if ($ReopenOnly) {
        $current = Read-Stats
        if (-not $current.connectionsSaved) { throw 'Connections did not survive reopening.' }
        Write-Host 'PASS saved workbook recalculates on open with persisted connections and native formulas.'
    } else {
        $research.Range('C7').Value2 = 'energy'
        $energy = Verify-Selection
        if ($energy -eq $baseline) { throw 'Subject filter did not change the result.' }
        $research.Range('C8').Value2 = [string]$research.Range('E19').Value2
        [void](Verify-Selection)
        if ($research.Range('F8').Value2 -ne 1) { throw 'Exact sender filter should retain one sender.' }
        if ($queries.Range('C6').Value2 -ne $root) { throw 'Text filters recomputed the S3 root.' }
        Write-Host 'PASS subject and sender edits update results while retaining the S3 table.'
        $research.Range('C7').Value2 = 'xlduck-no-match-8c66df2c'
        [void](Verify-Selection)
        if ($research.Range('B19').Value2 -or $research.Range('E19').Value2) { throw 'An empty selection left stale spilled rows.' }
        Write-Host 'PASS empty selection clears tables and leaves the reply share blank.'
        $research.Range('C7').Value2 = 'california'; $research.Range('C8').Value2 = 'All'
        $research.Range('C6').Value2 = 2000
        [void](Verify-Selection)
        if ($queries.Range('C6').Value2 -eq $root) { throw 'Changing year did not replace the S3 root.' }
        $research.Range('C6').Value2 = 2001
        [void](Verify-Selection)
        Write-Host 'PASS changing the year reloads the source and updates the dependent results.'
        $before = Read-Stats
        $retained = $queries.Range('C6').Value2
        $session.Application.CalculateFull()
        [void](Verify-Selection)
        if ($queries.Range('C6').Value2 -ne $retained -or (Read-Stats).statistics.Executions -ne $before.statistics.Executions) { throw 'Unchanged recalculation executed queries again.' }
        Write-Host 'PASS unchanged full recalculation reuses results.'
    }
    $snapshot = @{}
    foreach ($name in @('Research','Queries')) {
        $target = if ($name -eq 'Research') { $research } else { $queries }
        foreach ($formula in $contract.formulas.$name.PSObject.Properties) {
            if ($target.Range($formula.Name).Formula2.Contains('@')) { throw "Implicit intersection in $name!$($formula.Name)." }
        }
        $matrix = $target.Range($contract.ranges.$name).Value2
        $rows = @()
        for ($r=1; $r -le $matrix.GetLength(0); $r++) {
            $row = @()
            for ($c=1; $c -le $matrix.GetLength(1); $c++) {
                $v = $matrix[$r,$c]
                if (($v -is [int] -and $v -lt -2000000000) -or ($v -is [string] -and $v -match '^#(ERROR|NAME|REF|VALUE|SPILL|N/A)')) { throw "Formula error in $name at row $r, column $c." }
                $row += $v
            }
            $rows += ,$row
        }
        $snapshot[$name] = $rows
    }
    $snapshot | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $support 'calculated.json') -Encoding UTF8
    if ($Prepare) {
        # Only our explicitly owned verification sheet is removed, before handing over.
        $session.Application.DisplayAlerts = $false
        try { $check.Delete() } finally { $session.Application.DisplayAlerts = $true }
        $research.Activate(); $research.Range('C12').Select() | Out-Null
        $session.Application.ActiveWindow.Zoom = 90
        $session.Workbook.Save()
        Write-Host 'Saved native-calculated demo with all inputs restored and no verification sheet.'
    }
    Write-Host 'PASS native formulas, dynamic spills and final error scan.'
} catch {
    Write-Host ('Demo verification failed: ' + $_.Exception.Message)
    Write-Host $_.ScriptStackTrace
    throw
} finally {
    $target = $null
    if ($check) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($check) }
    if ($queries) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($queries) }
    if ($research) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($research) }
    if ($session) { Close-XlDuckExcelSession $session }
}
