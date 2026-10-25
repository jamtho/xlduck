# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([switch]$NoBuild, [string]$AddInPath, [string]$RuntimeRoot, [string]$FrameworkProbe)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
if (-not $AddInPath) {
    if (-not $NoBuild) {
        dotnet build (Join-Path $PSScriptRoot '../XlDuck') --nologo -m:1 /nodeReuse:false /p:UseSharedCompilation=false
        if ($LASTEXITCODE -ne 0) { throw 'Production build failed.' }
    }
    $AddInPath = Join-Path $PSScriptRoot '../XlDuck/bin/Debug/net8.0-windows/XlDuck-AddIn64.xll'
}
$privateBuild = Copy-XlDuckDevelopmentBuild $AddInPath
$session = $null; $sheet = $null; $other = $null; $otherSheet = $null
$closeGuard = $null
$caseDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) ('artifacts/production-tests/' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($caseDirectory)
$savedPaths = @()
$passed = 0
function Formula([string]$address, [string]$formula) { Invoke-XlDuckExcel { $sheet.Range($address).Formula2 = $formula } }
function Wait-Value([string]$address, $expected, [int]$seconds = 15) {
    [void](Wait-XlDuckCell $sheet $address -TimeoutSeconds $seconds -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -ceq $expected
    })
}
function Wait-Handle([string]$address) {
    Wait-XlDuckCell $sheet $address -TimeoutSeconds 15 -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -is [string] -and $value -cmatch '^xlduck://[tfp]/[1-9][0-9]*::[A-Za-z0-9_-]{11}$'
    }
}
function Command([string]$name) { Invoke-XlDuckExcel { $excel.Run($name, $session.Workbook.Name) } | Out-Null }
function Pass([string]$name) { $script:passed++; Write-Host "PASS $name" }
function Stats {
    Invoke-XlDuckExcel { $sheet.Range('Z1').Calculate() }
    $sheet.Range('Z1').Value2 | ConvertFrom-Json
}
function Wait-Status([scriptblock]$predicate) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $status = Stats
        if (& $predicate $status) { return $status }
        Start-Sleep -Milliseconds 30
    } while ($watch.Elapsed.TotalSeconds -lt 20)
    throw "Timed out waiting for workbook status: $($status | ConvertTo-Json -Compress)"
}
try {
    if ($RuntimeRoot) {
        $session = New-XlDuckExcelSession -AddInPath (Copy-XlDuckDevelopmentBuild $FrameworkProbe)
        $session.Application.Run('XlDuckProbeRuntimeRoot', [IO.Path]::GetFullPath($RuntimeRoot)) | Out-Null
        if (-not $session.Application.RegisterXLL($privateBuild)) { throw 'Production add-in could not load.' }
    } else { $session = New-XlDuckExcelSession -AddInPath $privateBuild }
    $excel = $session.Application; $sheet = $session.Workbook.Worksheets.Item(1)
    Formula 'A1' '=DuckQuery("SELECT 1 n")'
    Formula 'Z1' '=DuckStatus()'
    if ($RuntimeRoot -and ([version](Stats).runtime).Major -ne 10) { throw 'Portable .NET 10 runtime was not selected.' }
    Formula 'D1' '=DuckScalar(A1)'
    Wait-Value 'D1' 1
    if ((Wait-Handle 'A1').Length -gt 27) { throw 'Table handles are not compact numbered aliases.' }
    Pass 'new workbooks compute automatically and return compact handles'
    $sheet.Range('B1').Value2 = 2
    Formula 'A1' '=DuckQuery("SELECT ? n", B1)'
    for ($i = 2; $i -le 9; $i++) { Formula "A$i" ('=DuckQuery("SELECT n+1 n FROM ?", A{0})' -f ($i - 1)) }
    Formula 'D1' '=DuckScalar(A9)'
    Wait-Value 'D1' 10
    $watch = [Diagnostics.Stopwatch]::StartNew(); $sheet.Range('B1').Value2 = 10; Wait-Value 'D1' 18
    Pass "production worker chain updates ($([Math]::Round($watch.Elapsed.TotalMilliseconds)) ms)"
    $handle = Wait-Handle 'A1'; $executions = (Stats).statistics.Executions
    Invoke-XlDuckExcel { $excel.CalculateFull() }
    Formula 'A1' $sheet.Range('A1').Formula2
    Wait-Value 'D1' 18
    if ((Wait-Handle 'A1') -ne $handle -or (Stats).statistics.Executions -ne $executions) { throw 'Unchanged calculation reran SQL.' }
    Pass 'full calculation and unchanged formula reentry retain results'
    $sheet.Range('J40').Value2 = 'SELECT ? * ? AS n'
    $sqlReference = "='" + $sheet.Name.Replace("'", "''") + "'!" + '$J$40'
    $session.Workbook.Names.Add('Demo_Scale_SQL', $sqlReference) | Out-Null
    Formula 'J41' '=DuckQueryOutScalar(J40, 6, 7)'; Wait-Value 'J41' 42
    Formula 'J42' '=DuckQueryOutScalar(Demo_Scale_SQL, 9, 5)'; Wait-Value 'J42' 45
    Formula 'J43' '=DuckQuery(Demo_Scale_SQL, 2, 3)'
    Formula 'J44' '=DuckScalar(J43)'; Wait-Value 'J44' 6
    $sheet.Range('J40').Value2 = 'SELECT ? + ? AS n'
    Wait-Value 'J41' 13; Wait-Value 'J42' 14; Wait-Value 'J44' 5
    Pass 'SQL in ordinary and named cells supports repeated calls and updates its dependents'
    Formula 'F1' '=DuckQueryOutScalar("SELECT ?+?", 20, 22) + DuckQueryOutScalar("SELECT 1")'
    Wait-Value 'F1' 43
    Formula 'F2' '=DuckScalar(DuckQuery("SELECT 42"))'
    Wait-Value 'F2' 42
    $before = (Stats).statistics.Executions
    Invoke-XlDuckExcel { $excel.CalculateFull() }; Wait-Value 'F1' 43; Wait-Value 'F2' 42
    if ((Stats).statistics.Executions -ne $before) { throw 'Multiple calls in a formula lost their leases.' }
    Pass 'nested and multiple XlDuck calls in a cell keep independent live bindings'
    Formula 'G1' '=DuckQueryOut("SELECT ''00123'' text_value, 9007199254740993::HUGEINT exact_value, 1.234567890123456789::DECIMAL(20,18) price")'
    Wait-Value 'G2' '00123'; Wait-Value 'H2' '9007199254740993'; Wait-Value 'I2' '1.234567890123456789'
    Pass 'spilled output preserves text, large integers and exact decimals'
    Formula 'T1' '=DuckQueryOut("SELECT ? a, ? b, ? c, ''?'' question", "00123", "", 7)'
    Wait-Value 'T2' '00123'; Wait-Value 'U2' ''; Wait-Value 'V2' 7; Wait-Value 'W2' '?'
    Pass 'typed arguments preserve numeric text, empty strings and literal question marks'
    $sheet.Range('T5').Value2 = 'value'; $sheet.Range('T6').Value2 = 3; $sheet.Range('T7').Value2 = 4
    Formula 'T9' '=DuckCapture(T5:T7)'
    Formula 'T10' '=DuckFrag("SELECT value*2 v FROM ?", T9)'
    Formula 'T11' '=DuckQuery("SELECT sum(v) n FROM ?", T10)'
    Formula 'U11' '=DuckScalar(T11)'; Wait-Value 'U11' 14
    Formula 'T12' '=DuckInspect(T9,"rows")'; Wait-Value 'T12' 2
    Formula 'U14' '=DuckInspect(T11)'; Wait-Value 'U15' 'n'
    Formula 'T16' '=DuckPlot(T10,"bar","x","v","y","v")'; [void](Wait-Handle 'T16')
    Formula 'T17' '=DuckPlot(T10,"map","layer","satellite")'
    [void](Wait-XlDuckCell $sheet 'T17' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR: Unknown plot option: layer') })
    $sheet.Range('T17').ClearContents()
    $sheet.Range('T7').Value2 = 5; Wait-Value 'U11' 16
    Pass 'capture, lazy composition, inspection and plot definitions form a working research DAG'
    Formula 'T20' '=NA()'; Formula 'T21' '=DuckCapture(T20:T20)'
    [void](Wait-XlDuckCell $sheet 'T21' -Until { param($v) $v -is [string] -and $v.Contains('Capture contains an Excel error') })
    Pass 'capture rejects Excel errors in headers as well as data'
    $oldHandle = Wait-Handle 'A1'; $sheet.Name = "Renamed 'sheet"
    $sheet.Range('B1').Value2 = 11; Wait-Value 'D1' 19
    if ((Stats).ready -ne $true) { throw 'Sheet rename lost the session.' }
    Pass 'sheet rename preserves the workbook session and publication'
    foreach ($sql in @('CREATE TABLE forbidden(i INT)', 'COPY (SELECT 1) TO ''forbidden.csv''', 'SELECT * FROM enable_logging()')) {
        Formula 'F4' ('=DuckQuery("{0}")' -f $sql)
        [void](Wait-XlDuckCell $sheet 'F4' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
    }
    Formula 'F4' '=DuckQueryOutScalar("SELECT count(*) FROM duckdb_tables() WHERE table_name=''forbidden''")'; Wait-Value 'F4' 0
    Pass 'the production worksheet path rejects mutations without effects'
    foreach ($removed in @('DuckExecute', 'DuckConfig', 'DuckConfigReady', 'DuckQueryAfterConfig', 'GetLogPath')) {
        Formula 'F5' ("=$removed()")
        [void](Wait-XlDuckCell $sheet 'F5' -Until { param($v) $v -eq -2146826259 })
    }
    $sheet.Range('F5').ClearContents()
    Pass 'legacy write/configuration functions and public helpers are not registered'
    $other = $excel.Workbooks.Add(); $otherSheet = $other.Worksheets.Item(1)
    $otherSheet.Range('A1').Formula2 = '=DuckQuery("SELECT 444 n")'
    $otherHandle = Wait-XlDuckCell $otherSheet 'A1' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[1-9][0-9]*::[A-Za-z0-9_-]{11}$' }
    $sheet.Range('R2').Value2 = $otherHandle
    Formula 'F6' '=DuckScalar(R2)'
    [void](Wait-XlDuckCell $sheet 'F6' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('different workbook') })
    $sheet.Range('F6').ClearContents()
    $beforeRestart = Wait-Handle 'A1'; Command 'DuckRestartWorkbook'; Wait-Value 'D1' 19
    if ((Wait-Handle 'A1') -eq $beforeRestart -or $otherSheet.Range('A1').Value2 -ne $otherHandle) { throw 'Restart crossed workbook boundaries.' }
    $sheet.Range('R2').Value2 = $beforeRestart
    Formula 'F6' '=DuckScalar(R2)'
    [void](Wait-XlDuckCell $sheet 'F6' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('earlier session') })
    $sheet.Range('F6').ClearContents()
    Pass 'compact handles reject foreign and stale results; restart is isolated between workbooks'
    $session.Workbook.Activate()
    $session.Workbook.Date1904 = $true
    [void](Wait-Status { param($s) $s.statistics.Pending -eq 0 })
    Formula 'F8' '=DuckQueryOutScalar("SELECT DATE ''2024-01-01''")'
    # The lifecycle reconciliation observes date-system changes without rereading retained SQL results.
    Wait-Value 'F8' 43830
    Formula 'F9' '=DuckDate(0)'; Wait-Value 'F9' '1904-01-01'
    $session.Workbook.Date1904 = $false
    Wait-Value 'F8' 45292
    Pass 'date output and helpers respect workbook date systems'
    Command 'DuckPauseQueries'; $beforePause = (Stats).statistics.Executions
    Formula 'F11' '=DuckQueryOutScalar("SELECT 777")'
    if ((Stats).statistics.Executions -ne $beforePause) { throw 'A paused query started execution.' }
    Command 'DuckResumeQueries'; Wait-Value 'F11' 777
    Pass 'pause holds new work; resume publishes the retained request'
    Formula 'F13' '=DuckQueryOutScalar("SELECT sum(sqrt(i)) FROM range(1000000000) r(i)")'
    [void](Wait-Status { param($s) $s.statistics.Pending -gt 0 })
    Command 'DuckInterrupt'
    [void](Wait-XlDuckCell $sheet 'F13' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('cancel') })
    Formula 'F13' '=DuckQueryOutScalar("SELECT 888")'; Wait-Value 'F13' 888
    Pass 'cancel interrupts running work and the worker remains usable'
    $source = Join-Path $caseDirectory 'source.csv'; $sheet.Range('R1').Value2 = $source
    Formula 'F15' '=DuckQueryOutScalar("SELECT sum(n) FROM read_csv(?)", R1)'
    [void](Wait-XlDuckCell $sheet 'F15' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
    [IO.File]::WriteAllText($source, "n`n99`n")
    Formula 'F15' $sheet.Range('F15').Formula2
    if ($sheet.Range('F15').Value2 -isnot [string] -or -not $sheet.Range('F15').Value2.StartsWith('#ERROR:')) { throw 'Failure retried implicitly.' }
    Command 'DuckRetryFailed'; Wait-Value 'F15' 99
    Pass 'external read failures stay stable until an explicit retry'
    $excel.Calculation = -4135
    Formula 'F18' '=DuckQueryOutScalar("SELECT 314")'; $sheet.Range('F18').Calculate()
    [void](Wait-Status { param($s) $s.statistics.Pending -eq 0 })
    if ($excel.Calculation -ne -4135) { throw 'XlDuck changed Excel calculation mode.' }
    $sheet.Range('F18').Calculate(); Wait-Value 'F18' 314
    $manualSource = Join-Path $caseDirectory 'manual.csv'; $sheet.Range('R2').Value2 = $manualSource
    Formula 'F19' '=DuckQueryOutScalar("SELECT sum(n) FROM read_csv(?)", R2)'; $sheet.Range('F19').Calculate()
    [void](Wait-Status { param($s) $s.statistics.Pending -eq 0 }); $sheet.Range('F19').Calculate()
    [void](Wait-XlDuckCell $sheet 'F19' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
    [IO.File]::WriteAllText($manualSource, "n`n271`n")
    Command 'DuckRetryFailed'; $sheet.Range('F19').Calculate()
    [void](Wait-Status { param($s) $s.statistics.Pending -eq 0 }); $sheet.Range('F19').Calculate(); Wait-Value 'F19' 271
    $excel.Calculation = -4105
    Pass 'manual calculation remains manual and supports completion and delayed failure retry'
    $interop = Join-Path $env:USERPROFILE '.nuget/packages/exceldna.interop/15.0.1/lib/net452/Microsoft.Office.Interop.Excel.dll'
    Add-Type -Path $interop
    Add-Type -ReferencedAssemblies $interop -TypeDefinition @'
using System;
using Excel = Microsoft.Office.Interop.Excel;
public sealed class XlDuckCancelClose : IDisposable {
    private Excel.AppEvents_Event events;
    private string name;
    public XlDuckCancelClose(object application, string workbookName) {
        events = (Excel.AppEvents_Event)application; name = workbookName; events.WorkbookBeforeClose += Closing;
    }
    private void Closing(Excel.Workbook book, ref bool cancel) { if (book.Name == name) cancel = true; }
    public void Dispose() { events.WorkbookBeforeClose -= Closing; events = null; }
}
'@
    $closeGuard = New-Object XlDuckCancelClose($excel, $session.Workbook.Name)
    $beforeClose = (Stats).session
    $session.Workbook.Close($false)
    [void](Stats)
    if ((Stats).session -ne $beforeClose) { throw 'Cancelled close destroyed the session.' }
    $closeGuard.Dispose(); $closeGuard = $null
    $sheet.Range('B1').Value2 = 12; Wait-Value 'D1' 20
    Pass 'a cancelled workbook close preserves the session and subsequent updates'
    $beforeConnections = (Stats).session
    $connectionJson = '{"S3":[{"Scope":"s3://fixture/","Credentials":0,"Endpoint":"127.0.0.1:1","UseSsl":false,"PathStyle":true}]}'
    $excel.Run('DuckSetConnections', $connectionJson, $session.Workbook.Name) | Out-Null
    [void](Wait-Status { param($s) $s.connections -eq 1 -and $s.session -ne $beforeConnections })
    Wait-Value 'D1' 20
    if ($otherSheet.Range('A1').Value2 -ne $otherHandle) { throw 'Connection changes restarted another workbook.' }
    Pass 'connection settings apply without enablement and restart only their workbook'
    $saved = Join-Path $caseDirectory 'research.xlsx'; $savedPaths += $saved
    $session.Workbook.SaveAs($saved, 51)
    [void](Wait-Status { param($s) $s.connectionsSaved })
    $savedSession = (Stats).session
    $renamed = Join-Path $caseDirectory 'renamed research.xlsx'; $savedPaths += $renamed
    $session.Workbook.SaveAs($renamed, 51)
    [void](Wait-Status { param($s) $s.connectionsSaved })
    if ((Stats).session -ne $savedSession) { throw 'Save As restarted the live computation session.' }
    Pass 'Save As keeps live results and retains connection settings'
    [IO.File]::WriteAllText($source, "n`n100`n")
    $session.Workbook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Workbook)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null
    $session.Workbook = $excel.Workbooks.Open($renamed); $sheet = $session.Workbook.Worksheets.Item(1)
    [void](Wait-Status { param($s) $s.ready -and $s.connections -eq 1 -and $s.session -ne $savedSession })
    Wait-Value 'F15' 100; Wait-Value 'D1' 20
    Pass 'opening an ordinary saved workbook starts a fresh session and rereads its sources'
    $copy = Join-Path $caseDirectory 'unfamiliar copy.xlsx'
    $originalSession = (Stats).session
    [IO.File]::Copy($renamed, $copy)
    $copyBook = $excel.Workbooks.Open($copy); $copySheet = $copyBook.Worksheets.Item(1)
    try {
        [void](Wait-XlDuckCell $copySheet 'F15' -Until { param($v) $v -eq 100 })
        [void](Wait-XlDuckCell $copySheet 'Z1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).session -ne $originalSession -and ($v | ConvertFrom-Json).connections -eq 0 })
    } finally { $copyBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copySheet); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copyBook) }
    Pass 'an ordinary workbook copy computes automatically in its own session'
    $audit = $excel.Workbooks.Add(); $auditSheet = $audit.Worksheets.Item(1)
    try {
        $auditSheet.Range('Z1').Formula2 = '=DuckStatus()'
        $auditSheet.Range('A50').Formula2 = '=DuckQuery("SELECT 7654321")'
        [void](Wait-XlDuckCell $auditSheet 'A50' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[1-9][0-9]*::[A-Za-z0-9_-]{11}$' })
        $auditSheet.Rows.Item(45).Insert() | Out-Null
        [void](Wait-XlDuckCell $auditSheet 'A51' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[1-9][0-9]*::[A-Za-z0-9_-]{11}$' })
        $auditSheet.Range('A51').ClearContents() | Out-Null
        [void](Wait-XlDuckCell $auditSheet 'Z1' -Until {
            param($v)
            $auditSheet.Range('Z1').Calculate()
            ($v | ConvertFrom-Json).statistics.ReferencedNodes -eq 0
        })
        Pass 'moving and deleting a formula releases its old cell ownership'
    } finally { $audit.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($auditSheet); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($audit) }
    Write-Host "$passed production checks passed."
} catch {
    Write-Host "PRODUCTION FAILURE: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    throw
} finally {
    if ($closeGuard) { $closeGuard.Dispose() }
    if ($other) { Invoke-XlDuckExcel { $other.Close($false) }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($other) }
    if ($otherSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session) {
        try { Close-XlDuckExcelSession $session }
        catch { Write-Host $_.ScriptStackTrace; throw }
    }
    foreach ($savedPath in $savedPaths) {
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $key = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($savedPath).ToUpperInvariant()))).Replace('-', '') }
        finally { $sha.Dispose() }
        $connectionRecord = Join-Path $env:LOCALAPPDATA ("XlDuck/connections/$key.json")
        if (Test-Path -LiteralPath $connectionRecord) { Remove-Item -LiteralPath $connectionRecord }
    }
}
