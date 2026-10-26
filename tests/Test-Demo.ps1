# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$WorkbookPath, [string]$AddInPath, [switch]$Prepare)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/DemoSession.ps1')
$repo = Split-Path $PSScriptRoot -Parent
$contract = Get-Content (Join-Path $repo 'artifacts/demo-authoring/contract.json') -Raw | ConvertFrom-Json
$session = $null; $sheet = $null; $source = $null
function Stats {
    Invoke-XlDuckExcel { $sheet.Range('C47').Formula2 = '=DuckStatus()'; $sheet.Range('C47').Calculate() }
    $sheet.Range('C47').Value2 | ConvertFrom-Json
}
function Verify-Results([int]$day = 8, [string]$site = 'All', [double]$minimum = 35) {
    $expected = @($contract.rows | Where-Object { $_[4] -and $_[1] -ge $day -and ($site -eq 'All' -or $_[2] -eq $site) -and $_[3] -ge $minimum })
    [void](Wait-XlDuckCell $sheet 'F7' -Until { param($v) $v -ceq [double]$expected.Count })
    if ($expected.Count) {
        $mean = ($expected | ForEach-Object { [double]$_[3] } | Measure-Object -Average).Average
        [void](Wait-XlDuckCell $sheet 'F8' -Until { param($v) $v -is [double] -and [Math]::Abs($v - $mean) -lt 0.00000001 })
        $groups = @($expected | Group-Object { $_[2] } | Sort-Object Name)
        for ($i = 0; $i -lt $groups.Count; $i++) {
            $row = $i + 16; $group = $groups[$i]
            [void](Wait-XlDuckCell $sheet "E$row" -Until { param($v) $v -ceq $group.Name })
            [void](Wait-XlDuckCell $sheet "F$row" -Until { param($v) $v -ceq [double]$group.Count })
        }
    } else { [void](Wait-XlDuckCell $sheet 'F8' -Until { param($v) $null -eq $v -or $v -ceq '' }) }
    foreach ($cell in @('C11','C12','C16','C17','C18','C19','C20','C21')) {
        [void](Wait-XlDuckCell $sheet $cell -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[1-9][0-9]*::[A-Za-z0-9_-]{11}$' })
    }
    return $expected.Count
}
try {
    $session = New-XlDuckDemoSession -WorkbookPath $WorkbookPath -AddInPath $AddInPath -CopyWorkbook:(-not $Prepare)
    $sheet = $session.Workbook.Worksheets.Item('Research')
    $source = $session.Workbook.Worksheets.Item('Observations')
    if ($Prepare) {
        # The artifact renderer cannot register Excel-DNA UDFs or their spill metadata.
        # Write the authored formulas through Formula2 in the intended calculation engine.
        foreach ($property in $contract.formulas.PSObject.Properties) {
            Invoke-XlDuckExcel { $sheet.Range($property.Name).Formula2 = $property.Value }
        }
    }
    $baseline = Verify-Results
    if ($source.Range('B7').Value2 -cne '0001' -or $source.Range('B7').Text -cne '0001' -or $sheet.Range('B30').Value2 -cne '0022' -or $sheet.Range('B30').Text -cne '0022') { throw 'The synthetic sample identifier lost its leading zeros.' }
    if ($sheet.Range('C11').Text -cne $sheet.Range('C11').Value2) { throw 'The demo must display its actual chart handle.' }
    Write-Host "PASS default demo: $baseline observations, exact identifiers, summaries and plot handles."
    $sheet.Range('C7').Value2 = 15
    [void](Verify-Results 15)
    $sheet.Range('C8').Value2 = 'Ridge'
    [void](Verify-Results 15 'Ridge')
    $sheet.Range('C9').Value2 = 60
    [void](Verify-Results 15 'Ridge' 60)
    Write-Host 'PASS all three inputs update the query graph and outputs.'
    $sheet.Range('C9').Value2 = 100
    [void](Verify-Results 15 'Ridge' 100)
    if ($sheet.Range('E16').Value2) { throw 'An empty result retained an old summary row.' }
    Write-Host 'PASS empty selection clears spilled results and leaves the mean blank.'
    $sheet.Range('C7').Value2 = 8; $sheet.Range('C8').Value2 = 'All'; $sheet.Range('C9').Value2 = 35
    [void](Verify-Results)
    $original = $contract.rows[21][3]
    $contract.rows[21][3] = $original + 10
    $source.Range('E28').Value2 = [double]$contract.rows[21][3]
    [void](Verify-Results)
    $contract.rows[21][4] = $false; $source.Range('F28').Value2 = $false
    [void](Verify-Results)
    $contract.rows[21][3] = $original; $contract.rows[21][4] = $true
    $source.Range('E28').Value2 = [double]$original; $source.Range('F28').Value2 = $true
    [void](Verify-Results)
    Write-Host 'PASS edits to typed observations and accepted flags update the capture.'
    $retained = $sheet.Range('C19').Value2; $executions = (Stats).statistics.Executions
    $session.Application.CalculateFull(); [void](Verify-Results)
    if ($sheet.Range('C19').Value2 -ne $retained -or (Stats).statistics.Executions -ne $executions) { throw 'Unchanged demo calculation reread its data.' }
    Write-Host 'PASS unchanged full calculation reuses retained results.'
    $sheet.Range('C47').ClearContents()
    foreach ($cell in @('C11','C12','C16','C17','C18','C19','C20','C21','E15','E22','B29','F7','F8')) {
        if ($sheet.Range($cell).Formula2.Contains('@')) { throw "Implicit intersection broke $cell." }
    }
    $values = $sheet.Range('B1:G47').Value2
    $rows = @()
    for ($r = 1; $r -le 47; $r++) {
        $cells = @()
        for ($c = 1; $c -le 6; $c++) {
            $value = $values[$r,$c]
            if ($value -is [string] -and $value -match '^#(ERROR|BLOCKED|NAME|REF|VALUE|SPILL|N/A)') { throw "Demo contains an error at row $r, column $c." }
            if ($value -is [int] -and $value -lt -2000000000) { throw "Demo contains an Excel error at row $r, column $c." }
            $cells += $value
        }
        $rows += ,$cells
    }
    @{research=$rows} | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $repo 'artifacts/demo-authoring/calculated.json') -Encoding UTF8
    if ($Prepare) {
        $sheet.Activate(); $sheet.Range('C11').Select() | Out-Null
        $session.Workbook.Save()
        Write-Host 'Saved native-calculated demo, with all test edits restored.'
    }
    Write-Host 'PASS demo formulas, dynamic spills and final error scan.'
} finally {
    if ($source) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($source) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session) { Close-XlDuckExcelSession $session }
}
