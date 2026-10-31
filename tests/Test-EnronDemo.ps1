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
function Read-Value($Sheet, [string]$Cell) { Invoke-XlDuckExcel { $Sheet.Range($Cell).Value2 } }
function Write-Value($Sheet, [string]$Cell, $Value) {
    # Separate COM setter sites: Windows PowerShell caches the first value's type.
    if ($Value -is [ValueType]) { Invoke-XlDuckExcel { $Sheet.Range($Cell).Value2 = [double]$Value } }
    else { Invoke-XlDuckExcel { $Sheet.Range($Cell).Value2 = [string]$Value } }
}
function Write-Formula($Sheet, [string]$Cell, [string]$Formula) { Invoke-XlDuckExcel { $Sheet.Range($Cell).Formula2 = $Formula } }
function Wait-Number([string]$Cell, [double]$Expected) {
    [void](Wait-XlDuckCell $research $Cell -Until { param($v) $v -is [double] -and [Math]::Abs($v - $Expected) -lt 1e-10 } -TimeoutSeconds 60)
}
function Read-Stats {
    Invoke-XlDuckExcel { $check.Range('K1').Formula2 = '=DuckStatus()'; $check.Range('K1').Calculate() }
    Read-Value $check 'K1' | ConvertFrom-Json
}
function Verify-Selection {
    # An independent scan of the source verifies the displayed aggregate and filters.
    Write-Value $check 'A1' ([string](Read-Value $research 'C5'))
    Write-Value $check 'A2' ([double](Read-Value $research 'C6'))
    Write-Value $check 'A3' ([string](Read-Value $research 'C7'))
    Write-Value $check 'A4' ([string](Read-Value $research 'C8'))
    Write-Value $check 'A5' @'
SELECT count(*) AS records,
       count(DISTINCT lower(coalesce(nullif(trim(from_email), ''), nullif(trim(from_name), ''), '(unknown)'))) AS senders,
       avg(coalesce(is_reply, false)::INTEGER) AS reply_share
FROM read_parquet(?)
WHERE year(timezone('UTC', date)) = ?
  AND (? = 'All' OR contains(lower(coalesce(subject_clean, subject, '')), lower(?)))
  AND (? = 'All' OR lower(coalesce(nullif(trim(from_email), ''), nullif(trim(from_name), ''), '(unknown)')) = lower(?))
'@
    Write-Formula $check 'D1' '=DuckQueryOut(A5,A1,A2,A3,A3,A4,A4)'
    [void](Wait-XlDuckCell $check 'D2' -Until { param($v) $v -is [double] } -TimeoutSeconds 60)
    $expected = [double](Read-Value $check 'D2')
    Wait-Number 'F7' $expected
    Wait-Number 'F8' ([double](Read-Value $check 'E2'))
    if ($expected -gt 0) { Wait-Number 'F9' ([double](Read-Value $check 'F2')) }
    else { [void](Wait-XlDuckCell $research 'F9' -Until { param($v) $null -eq $v -or $v -ceq '' }) }
    # Reconcile all monthly rows to the same count, and verify both live plot handles.
    Write-Formula $check 'K2' '=SUM(Research!C19:C30)'
    [void](Wait-XlDuckCell $check 'K2' -Until { param($v) $v -is [double] -and $v -eq $expected })
    foreach ($cell in @('C12','C13')) {
        Write-Formula $check 'K3' ('=DuckInspect(Research!' + $cell + ',"rows")')
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
    if (-not $ReopenOnly) { Invoke-XlDuckExcel { $session.Application.Run('DuckSetConnections',$connection,$session.Workbook.Name) } | Out-Null }
    if ($Prepare) {
        foreach ($tab in $contract.formulas.PSObject.Properties) {
            $target = if ($tab.Name -eq 'Queries') { $queries } else { $research }
            foreach ($formula in $tab.Value.PSObject.Properties) { Invoke-XlDuckExcel { $target.Range($formula.Name).Formula2 = $formula.Value } }
        }
    }
    $baseline = Verify-Selection
    if ($baseline -le 0) { throw 'Expected records for the default research selection.' }
    $root = Read-Value $queries 'C6'
    [void](Wait-XlDuckCell $research 'F11' -Until { param($v) $v -is [double] -and $v -gt 0 })
    Write-Host "PASS live Enron default: $(Read-Value $research 'F11') loaded records; $baseline matching records; monthly totals, sender count, reply share and both plots verified."
    if ($ReopenOnly) {
        $current = Read-Stats
        if (-not $current.connectionsSaved) { throw 'Connections did not survive reopening.' }
        Write-Host 'PASS saved workbook recalculates on open with persisted connections and native formulas.'
    } else {
        Write-Value $research 'C7' 'energy'
        $energy = Verify-Selection
        if ($energy -eq $baseline) { throw 'Subject filter did not change the result.' }
        Write-Value $research 'C8' ([string](Read-Value $research 'E19'))
        [void](Verify-Selection)
        if ((Read-Value $research 'F8') -ne 1) { throw 'Exact sender filter should retain one sender.' }
        if ((Read-Value $queries 'C6') -ne $root) { throw 'Text filters recomputed the S3 root.' }
        Write-Host 'PASS subject and sender edits update results while retaining the S3 table.'
        Write-Value $research 'C7' 'xlduck-no-match-8c66df2c'
        [void](Verify-Selection)
        if ((Read-Value $research 'B19') -or (Read-Value $research 'E19')) { throw 'An empty selection left stale spilled rows.' }
        Write-Host 'PASS empty selection clears tables and leaves the reply share blank.'
        Write-Value $research 'C7' 'california'; Write-Value $research 'C8' 'All'
        Write-Value $research 'C6' 2000
        [void](Verify-Selection)
        if ((Read-Value $queries 'C6') -eq $root) { throw 'Changing year did not replace the S3 root.' }
        Write-Value $research 'C6' 2001
        [void](Verify-Selection)
        Write-Host 'PASS changing the year reloads the source and updates the dependent results.'
        $before = Read-Stats
        $retained = Read-Value $queries 'C6'
        Invoke-XlDuckExcel { $session.Application.CalculateFull() }
        [void](Verify-Selection)
        if ((Read-Value $queries 'C6') -ne $retained -or (Read-Stats).statistics.Executions -ne $before.statistics.Executions) { throw 'Unchanged recalculation executed queries again.' }
        Write-Host 'PASS unchanged full recalculation reuses results.'
    }
    $snapshot = @{}
    foreach ($name in @('Research','Queries')) {
        $target = if ($name -eq 'Research') { $research } else { $queries }
        foreach ($formula in $contract.formulas.$name.PSObject.Properties) {
            if ((Invoke-XlDuckExcel { $target.Range($formula.Name).Formula2 }).Contains('@')) { throw "Implicit intersection in $name!$($formula.Name)." }
        }
        # Preserve the two-dimensional COM array through PowerShell's pipeline.
        $matrix = (Invoke-XlDuckExcel { [pscustomobject]@{Values = $target.Range($contract.ranges.$name).Value2} }).Values
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
        Invoke-XlDuckExcel { $session.Application.DisplayAlerts = $false }
        try { Invoke-XlDuckExcel { $check.Delete() } } finally { Invoke-XlDuckExcel { $session.Application.DisplayAlerts = $true } }
        Invoke-XlDuckExcel { $research.Activate(); $research.Range('C12').Select() } | Out-Null
        Invoke-XlDuckExcel { $session.Application.ActiveWindow.Zoom = 90 }
        Invoke-XlDuckExcel { $session.Workbook.Save() }
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
