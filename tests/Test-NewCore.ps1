# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([switch]$NoBuild)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
if (-not $NoBuild) {
    foreach ($project in @('XlDuck.Worker/XlDuck.Worker.csproj', 'tests/XlDuck.TestAddIn/XlDuck.TestAddIn.csproj')) {
        dotnet build (Join-Path $PSScriptRoot "../$project") --nologo -m:1 /nodeReuse:false /p:UseSharedCompilation=false
        if ($LASTEXITCODE -ne 0) { throw "Build failed: $project" }
    }
}
$session = $null; $sheet = $null; $otherBook = $null; $otherSheet = $null
$passed = 0
function Formula([string]$address, [string]$formula) { Invoke-XlDuckExcel { $sheet.Range($address).Formula2 = $formula } }
function Wait-Value([string]$address, $expected, [int]$seconds = 10) {
    [void](Wait-XlDuckCell $sheet $address -TimeoutSeconds $seconds -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -ceq $expected
    })
}
function Wait-Handle([string]$address) {
    Wait-XlDuckCell $sheet $address -TimeoutSeconds 10 -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -is [string] -and $value.StartsWith('xlduck://')
    }
}
function Stats { $excel.Run('CoreStatistics', $session.WorkbookName) | ConvertFrom-Json }
function Pass([string]$name) { $script:passed++; Write-Host "PASS $name" }
try {
    $session = New-XlDuckExcelSession -AddInPath (Join-Path $PSScriptRoot 'XlDuck.TestAddIn/bin/Debug/net8.0-windows/XlDuck.TestAddIn-AddIn64.xll')
    $excel = $session.Application; $sheet = $session.Workbook.Worksheets.Item(1)
    $worker = (Resolve-Path (Join-Path $PSScriptRoot '../XlDuck.Worker/bin/Debug/net8.0/XlDuck.Worker.exe')).Path
    $excel.Run('CoreInitialize', $worker, $session.WorkbookName) | Out-Null
    Invoke-XlDuckExcel { $sheet.Range('B1').Value2 = 1 }
    Formula 'A1' '=CoreQuery("SELECT ? AS value", B1)'
    for ($i = 2; $i -le 9; $i++) { Formula "A$i" ('=CoreQuery("SELECT value+1 AS value FROM ?", A{0})' -f ($i - 1)) }
    Formula 'D1' '=CoreScalar(A9)'
    Wait-Value 'D1' 9
    $watch = [Diagnostics.Stopwatch]::StartNew()
    Invoke-XlDuckExcel { $sheet.Range('B1').Value2 = 10 }
    Wait-Value 'D1' 18
    $chainMs = $watch.Elapsed.TotalMilliseconds
    Pass "typed worker chain and changed input ($([Math]::Round($chainMs)) ms)"
    $handle = $sheet.Range('A1').Value2
    $executions = (Stats).Executions
    Invoke-XlDuckExcel { $excel.CalculateFull() }
    Formula 'A1' $sheet.Range('A1').Formula2
    Wait-Value 'D1' 18
    if ($sheet.Range('A1').Value2 -ne $handle -or (Stats).Executions -ne $executions) { throw 'Unchanged calculation/reentry reran SQL.' }
    Pass 'full calculation and formula reentry reuse live results with zero historical cache'

    Formula 'F1' '=CoreQuery("SELECT ? AS text_value, ? AS empty_value, ? AS n", "00123", "", 7)'
    Formula 'G1' '=CoreOut(F1)'
    Wait-Value 'G2' '00123'; Wait-Value 'I2' 7
    if ($sheet.Range('G2').Value2 -isnot [string]) { throw 'Numeric-looking text changed type.' }
    Pass 'typed text and empty arguments survive Excel and worker transport'
    Formula 'F4' '=CoreQuery("SELECT 9007199254740993::HUGEINT AS exact_value")'
    Formula 'G4' '=CoreScalar(F4)'
    Wait-Value 'G4' '9007199254740993'
    if ($sheet.Range('G4').Value2 -isnot [string]) { throw 'Large integer was rounded into an Excel number.' }
    Pass 'large integer is returned as exact text'

    Invoke-XlDuckExcel { $sheet.Range('K1').Value2 = 'a'; $sheet.Range('L1').Value2 = 'a'; $sheet.Range('M1').Value2 = 'a_2' }
    Invoke-XlDuckExcel { $sheet.Range('K2').Value2 = 1; $sheet.Range('L2').Value2 = 2; $sheet.Range('M2').Value2 = 3 }
    Formula 'J4' '=CoreCapture(K1:M2)'; Formula 'K4' '=CoreOut(J4)'
    Wait-Value 'L4' 'a_3'; Wait-Value 'M5' 3
    Pass 'capture creates deterministic unique headers and typed output'

    Formula 'F7' '=CoreFragment("SELECT 1 AS x, 2 AS y")'
    Formula 'F8' '=CorePlot(F7,"point")'
    Formula 'G8' '=CorePreview(F8)'
    $preview = Wait-XlDuckCell $sheet 'G8' -TimeoutSeconds 10 -Until { param($v) $v -is [string] -and $v.StartsWith('{') }
    $page = $preview | ConvertFrom-Json
    if ($page.TotalRows -ne 1 -or $page.Rows[0][1].Text -ne '2') { throw 'Fragment plot data is incorrect.' }
    $beforePreview = (Stats).Executions
    Invoke-XlDuckExcel { $sheet.Range('G8').Calculate() }
    if ((Stats).Executions -ne $beforePreview) { throw 'Preview reread the fragment query.' }
    Pass 'fragment plots retain materialized preview data'

    $otherBook = $excel.Workbooks.Add(); $otherSheet = $otherBook.Worksheets.Item(1)
    $excel.Run('CoreInitialize', $worker, $otherBook.Name) | Out-Null
    Invoke-XlDuckExcel { $otherSheet.Range('A1').Formula2 = '=CoreQuery("SELECT 10 AS value")' }
    $otherHandle = Wait-XlDuckCell $otherSheet 'A1' -TimeoutSeconds 10 -Until { param($v) $v -is [string] -and $v.StartsWith('xlduck://') }
    $excel.Run('CoreRestart', $session.WorkbookName) | Out-Null
    Wait-Value 'D1' 18
    if ($sheet.Range('A1').Value2 -eq $handle) { throw 'Restart retained the old session.' }
    Invoke-XlDuckExcel { $otherSheet.Range('A1').Calculate() }
    if ($otherSheet.Range('A1').Value2 -ne $otherHandle) { throw 'Restart changed another workbook.' }
    Pass 'restart discards one workbook session and preserves another'

    Invoke-XlDuckExcel { $sheet.Range('A1:M20').ClearContents() }
    $watch.Restart()
    do {
        $state = Stats
        if ($state.Nodes -eq 0 -and $state.Pending -eq 0 -and $state.Retiring -eq 0) { break }
        Start-Sleep -Milliseconds 25
    } while ($watch.Elapsed.TotalSeconds -lt 10)
    if ($state.Nodes -ne 0 -or $state.Pending -ne 0 -or $state.Retiring -ne 0) { throw "Cleared formulas leaked state: $($state | ConvertTo-Json -Compress)" }
    Pass 'clearing formulas releases every live result and queued cleanup'
    Formula 'A1' '=CoreQuery("SELECT sum(sin(i::DOUBLE)) FROM range(1000000000) t(i)")'
    $watch.Restart()
    do { $state = Stats; if ($state.Pending -gt 0) { break }; Start-Sleep -Milliseconds 25 } while ($watch.Elapsed.TotalSeconds -lt 5)
    if ($state.Pending -eq 0) { throw 'Slow query did not start.' }
    Invoke-XlDuckExcel { $sheet.Range('A1').ClearContents() }
    Formula 'A2' '=CoreQuery("SELECT 42 AS answer")'; Formula 'D2' '=CoreScalar(A2)'
    Wait-Value 'D2' 42
    if ($null -ne $sheet.Range('A1').Value2) { throw 'Late completion changed a deleted cell.' }
    Pass 'deleting a pending formula cancels work and allows recovery'
    Invoke-XlDuckExcel { $sheet.Range('A1:M20').ClearContents() }
    $sourcePath = (Resolve-Path (Join-Path $PSScriptRoot 'data/sample.csv')).Path
    Invoke-XlDuckExcel { $sheet.Range('B1').Value2 = $sourcePath }
    Formula 'A1' '=CoreQuery("SELECT sum(value) FROM read_csv(?)", B1)'
    [void](Wait-XlDuckCell $sheet 'A1' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('disabled') })
    $settings = @{ EnableExternalReads = $true; S3 = @() } | ConvertTo-Json -Compress
    $excel.Run('CoreInitializeWithSettings', $worker, $session.WorkbookName, $settings) | Out-Null
    Formula 'A1' '=CoreQuery("SELECT sum(value) FROM read_csv(?)", B1)'; Formula 'D1' '=CoreScalar(A1)'
    Wait-Value 'D1' 600
    $excel.Run('CoreRestart', $session.WorkbookName) | Out-Null
    Wait-Value 'D1' 600
    Pass 'host enablement permits local reads and survives workbook restart'
    foreach ($formula in @('=CoreQuery("CREATE TABLE forbidden(i INT)")', '=CoreQuery("SELECT * FROM enable_logging()")', '=CoreFragment("SELECT * FROM query(''SELECT 1'')")')) {
        Formula 'F1' $formula
        [void](Wait-XlDuckCell $sheet 'F1' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('read-and-compute') })
    }
    Formula 'F2' '=CoreQuery("SELECT count(*) FROM duckdb_tables() WHERE table_name=''forbidden''")'; Formula 'G2' '=CoreScalar(F2)'
    Wait-Value 'G2' 0
    Pass 'worksheet queries and fragments reject mutation and dynamic SQL without effects'
    Write-Host "$passed Excel core checks passed."
} finally {
    if ($session -and $session.Application) {
        try { $session.Application.Run('CoreShutdown') | Out-Null } catch { Write-Warning $_ }
    }
    if ($otherBook) { $otherBook.Close($false) }
    if ($otherSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet); $otherSheet = $null }
    if ($otherBook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook); $otherBook = $null }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null }
    $excel = $null
    if ($session) { Close-XlDuckExcelSession $session }
}
