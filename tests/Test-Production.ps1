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
$metadata = $null; $migrationBook = $null; $migrationSheet = $null; $bindingRecords = @()
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
        $value -is [string] -and $value -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$'
    }
}
function Command([string]$name) { Invoke-XlDuckExcel { $excel.Run($name, $session.Workbook.Name) } | Out-Null }
function Pass([string]$name) { $script:passed++; Write-Host "PASS $name" }
function Cell-Status([string]$address) {
    $json = Invoke-XlDuckExcel { $sheet.Evaluate("DuckCellStatus($address)") }
    if ($json -isnot [string] -or -not $json.StartsWith('{')) { throw "Cell status did not return JSON: $json" }
    $json | ConvertFrom-Json
}
function Help-Text([string]$formula) {
    # Evaluate can wrap a scalar UDF result in a single-cell array.
    $values = @(Invoke-XlDuckExcel { $sheet.Evaluate($formula) })
    if ($values.Count -ne 1 -or $values[0] -isnot [string]) { throw 'Help did not return exactly one text value.' }
    [string]$values[0]
}
function Explore-Raw($request) {
    $json = $request | ConvertTo-Json -Depth 8 -Compress
    $formula = 'DuckExplore(AZ1000,"' + $json.Replace('"','""') + '")'
    $result = Invoke-XlDuckExcel { $sheet.Evaluate($formula) }
    if ($result -isnot [string] -or -not $result.StartsWith('{')) { throw "Exploration did not return JSON: $result" }
    $result | ConvertFrom-Json
}
function Explore($request) {
    $json = $request | ConvertTo-Json -Depth 8 -Compress
    if (('DuckExplore(AZ1000,"' + $json.Replace('"','""') + '")').Length -le 255) { return Explore-Raw $request }
    # Small synthetic-test chunks also exercise the native Evaluate limit; production helper packs dynamically.
    $count = [int][Math]::Ceiling($json.Length / 24.0)
    for ($index = 0; $index -lt $count; $index++) {
        $part = $json.Substring($index * 24, [Math]::Min(24, $json.Length - $index * 24))
        $ack = Explore-Raw @{op='stage';id=$request.id;index=$index;count=$count;text=$part}
        if ($ack.state -ne 'staged') { return $ack }
    }
    Explore-Raw @{op='submit';id=$request.id}
}
function Wait-Explore($request) {
    $result = Explore $request
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($result.state -notin @('succeeded','failed','cancelled') -and $watch.Elapsed.TotalSeconds -lt 15) {
        Start-Sleep -Milliseconds 30
        $result = Explore @{op='poll';id=$request.id}
    }
    if ($result.state -ne 'succeeded') { throw "Exploration failed: $($result | ConvertTo-Json -Depth 8 -Compress)" }
    $result
}
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
    $expectedVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path (Split-Path $privateBuild -Parent) 'XlDuck.dll')).ProductVersion.Split('+')[0]
    if ((Help-Text 'DuckVersion()') -ne $expectedVersion) { throw 'DuckVersion does not match the loaded binary version.' }
    Formula 'A1' '=DuckQuery("SELECT 1 n")'
    Formula 'Z1' '=DuckStatus()'
    if ($RuntimeRoot -and ([version](Stats).runtime).Major -ne 10) { throw 'Portable .NET 10 runtime was not selected.' }
    Formula 'D1' '=DuckScalar(A1)'
    Wait-Value 'D1' 1
    if ((Wait-Handle 'A1').Length -gt 27) { throw 'Table handles are not compact numbered aliases.' }
    Pass 'new workbooks compute automatically and return compact handles'
    Formula 'AY1' '=DuckConnections()'
    $emptyConnections = Wait-XlDuckCell $sheet 'AY1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).state -eq 'succeeded' }
    $emptyConnections = $emptyConnections | ConvertFrom-Json
    if ($emptyConnections.total -ne 0 -or $emptyConnections.connections.Count -ne 0) { throw 'A new workbook inherited connections.' }
    Pass 'connection discovery works without an anchor in a worksheet and reports an empty inventory'
    Formula 'AY2' '=DuckHelp()'
    Formula 'AY3' '=DuckVersion()'
    $beforeHelp = (Stats).statistics.Executions
    foreach ($topic in @('index','workflow','reference','officejs')) {
        $help = Help-Text ('DuckHelp("' + $topic + '")')
        $version = Help-Text 'DuckVersion()'
        $prefix = if ($topic -eq 'officejs') { "// XlDuck $version help:" } else { "XlDuck $version help:" }
        if ($help -isnot [string] -or -not $help.StartsWith($prefix) -or $help.Length -gt 30000) { throw "Packaged help unavailable: $topic; returned $($help | ConvertTo-Json -Compress -Depth 3)" }
        $expectedHelp = switch ($topic) { 'index' { 'DuckHelp("workflow")' } 'workflow' { 'DuckSchema' } 'reference' { 'DuckExplore' } 'officejs' { 'function xlduckReadJson' } }
        if (-not $help.Contains($expectedHelp)) { throw "Packaged help content incomplete: $topic" }
    }
    Formula 'AF1' '=DuckHelp()'
    [void](Wait-XlDuckCell $sheet 'AF1' -Until { param($v) $v -is [string] -and $v.Contains('DuckHelp("workflow")') })
    if (-not (Help-Text 'DuckHelp("missing")').StartsWith('#ERROR: Unknown help topic.')) { throw 'Unknown help topic was not actionable.' }
    if (-not (Help-Text 'DuckHelp(42)').StartsWith('#ERROR: Help topic must be text.')) { throw 'Invalid help topic was accepted.' }
    if ((Help-Text 'DuckHelp(" WORKFLOW ")') -ne (Help-Text 'DuckHelp("workflow")')) { throw 'Help topic normalization failed.' }
    if ((Stats).statistics.Executions -ne $beforeHelp) { throw 'Reading help submitted a query.' }
    $sheet.Range('AF1').ClearContents()
    Pass 'embedded help is available through formulas and Evaluate without running queries'
    $cellStatus = Cell-Status 'A1'
    if ($cellStatus.state -ne 'succeeded' -or -not $cellStatus.resultReady -or $cellStatus.computations[0].computation.rows -ne 1) { throw "Incorrect cell status: $($cellStatus | ConvertTo-Json -Depth 8 -Compress)" }
    Formula 'AF1' '=DuckCellStatus(A1)'
    [void](Wait-XlDuckCell $sheet 'AF1' -Until { param($v) $v -is [string] -and ($v | ConvertFrom-Json).state -eq 'succeeded' })
    Formula 'AF3' '=DuckSchema("SELECT error(''must not execute'') AS n, 1::BIGINT AS id")'
    Wait-Value 'AF4' 'n'; Wait-Value 'AG5' 'BIGINT'
    Formula 'AF7' '=DuckQuery("SELECT missing FROM (SELECT 1 n)")'
    [void](Wait-XlDuckCell $sheet 'AF7' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
    if ((Cell-Status 'AF7').failure.code -ne 'sql_binding') { throw 'Binding failure lost its category.' }
    Formula 'AF7' '=DuckQuery("SELECT ?", 1, 2)'
    [void](Wait-XlDuckCell $sheet 'AF7' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
    if ((Cell-Status 'AF7').failure.code -ne 'invalid_input') { throw 'Synchronous argument failure was not tracked.' }
    $sheet.Range('AF9').Value2 = 'blocked'
    Formula 'AF8' '=DuckQueryOut("SELECT 123 n")'
    [void](Wait-XlDuckCell $sheet 'AF8' -Until { param($v) $v -eq -2146826243 })
    $spill = Cell-Status 'AF8'
    if ($spill.failure.code -ne 'excel_spill' -or -not $spill.resultReady) { throw "Spill failure was mistaken for execution failure: $($spill | ConvertTo-Json -Depth 8 -Compress)" }
    $sheet.Range('AF1:AG12').ClearContents()
    Pass 'per-cell status works through Evaluate, separates failures and detects blocked spill; schema does not evaluate rows'
    $beforeExplore = [string]$sheet.UsedRange.Address()
    $exploration = Wait-Explore @{op='query';id='native-query';sql='SELECT unnest([1,2,3]) n';maxRows=2}
    if ($exploration.result.rows.Count -ne 2 -or -not $exploration.result.hasMore) { throw 'Exploration did not bound its rows.' }
    $exploration = Wait-Explore @{op='schema';id='native-schema';sql='SELECT error(''must not execute'') AS n, 1::BIGINT AS id'}
    if ($exploration.result.columns[1].type -ne 'BIGINT' -or $exploration.result.rows.Count -ne 0) { throw 'Exploration did not return schema only.' }
    $exploration = Wait-Explore @{op='inspect';id='native-inspect';handle=(Wait-Handle 'A1')}
    if ($exploration.computation.rows -ne 1) { throw 'Exploration did not inspect the retained handle.' }
    $exploration = Wait-Explore @{op='query';id='native-reference';sql='SELECT n + ? AS n FROM ?';arguments=@(41,(Wait-Handle 'A1'))}
    if ($exploration.result.rows[0][0].text -ne '42') { throw 'Exploration did not bind a scalar and retained handle.' }
    if ([string]$sheet.UsedRange.Address() -ne $beforeExplore -or $sheet.Range('AZ1000').Value2) { throw 'Exploration changed the worksheet.' }
    Pass 'bounded query, schema and retained-handle inspection work through Evaluate without worksheet edits'
    $longRequest = @{op='query';id='long-query';sql=('SELECT ? AS n /*' + ('long SQL ' * 100) + '*/');arguments=@(42)}
    $exploration = Wait-Explore $longRequest
    if ($exploration.result.rows[0][0].text -ne '42') { throw 'Long query transfer changed the query or arguments.' }
    $beforeResubmit = (Stats).statistics.Executions
    $again = Wait-Explore $longRequest
    if ($again.computation.handle -ne $exploration.computation.handle -or (Stats).statistics.Executions -ne $beforeResubmit) { throw 'Resubmitting a long request executed it again.' }
    $duplicateSubmit = Explore-Raw @{op='submit';id='long-query'}
    if ($duplicateSubmit.state -ne 'succeeded') { throw 'Duplicate submit was not idempotent.' }
    $longRequest.sql = $longRequest.sql.Replace('AS n','+ 1 AS n')
    if ((Explore $longRequest).state -ne 'failed') { throw 'Changed inputs with a staged request id were accepted.' }
    Pass 'requests longer than Evaluate limits transfer without cells and remain idempotent'
    Formula 'AF1' '=DuckFrag("SELECT error(''must not execute'') AS n")'
    [void](Wait-Handle 'AF1')
    Formula 'AF3' '=DuckInspect(AF1)'; Wait-Value 'AF4' 'n'
    $sheet.Range('AF1:AG6').ClearContents()
    Pass 'fragment schema inspection avoids materializing data expressions'
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
    $otherHandle = Wait-XlDuckCell $otherSheet 'A1' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$' }
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
    if ((Cell-Status 'F11').state -ne 'paused') { throw 'Pending cell did not report paused state.' }
    Formula 'AF1' '=DuckQuery("SELECT ? n", F11)'
    if ((Cell-Status 'AF1').state -ne 'waiting_inputs') { throw 'Dependent formula did not report waiting for inputs.' }
    if ((Stats).statistics.Executions -ne $beforePause) { throw 'A paused query started execution.' }
    Command 'DuckResumeQueries'; Wait-Value 'F11' 777
    [void](Wait-Handle 'AF1'); $sheet.Range('AF1').ClearContents()
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
    if ((Cell-Status 'F15').failure.code -ne 'source_access') { throw 'Source failure lost its category.' }
    [IO.File]::WriteAllText($source, "n`n99`n")
    Formula 'F15' $sheet.Range('F15').Formula2
    if ($sheet.Range('F15').Value2 -isnot [string] -or -not $sheet.Range('F15').Value2.StartsWith('#ERROR:')) { throw 'Failure retried implicitly.' }
    Command 'DuckRetryFailed'; Wait-Value 'F15' 99
    Pass 'external read failures stay stable until an explicit retry'
    $excel.Calculation = -4135
    Command 'DuckPauseQueries'
    Formula 'F18' '=DuckQueryOutScalar("SELECT 314")'; $sheet.Range('F18').Calculate()
    Command 'DuckResumeQueries'
    [void](Wait-Status { param($s) $s.statistics.Pending -eq 0 })
    if ($excel.Calculation -ne -4135) { throw 'XlDuck changed Excel calculation mode.' }
    $manualStatus = Cell-Status 'F18'
    if ($manualStatus.state -ne 'awaiting_calculation' -or -not $manualStatus.resultReady -or $manualStatus.calculationMode -ne 'manual') { throw "Incorrect manual completion status: $($manualStatus | ConvertTo-Json -Depth 8 -Compress)" }
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
    $beforeDiscovery = Stats
    $beforeDiscoveryRange = [string]$sheet.UsedRange.Address()
    $other.Activate()
    $connections = (Help-Text 'DuckConnections(AZ1000)') | ConvertFrom-Json
    if ($connections.state -ne 'succeeded' -or $connections.total -ne 1 -or $connections.connections[0].scope -ne 's3://fixture/' -or $connections.connections[0].endpoint -ne '127.0.0.1:1' -or $connections.connections[0].authentication -ne 'anonymous') { throw 'Anchored connection discovery returned the wrong workbook or metadata.' }
    $elsewhere = @(Invoke-XlDuckExcel { $otherSheet.Evaluate('DuckConnections(AZ1000)') })[0] | ConvertFrom-Json
    if ($elsewhere.state -ne 'succeeded' -or $elsewhere.total -ne 0) { throw 'Connection discovery crossed workbook boundaries.' }
    $fromCell = $sheet.Range('AY1').Value2 | ConvertFrom-Json
    if ($fromCell.total -ne 1 -or $fromCell.connections[0].scope -ne 's3://fixture/') { throw 'Changing connections did not refresh the discovery formula.' }
    foreach ($invalid in @('DuckConnections(42)', 'DuckConnections(AZ1000:AZ1001)', 'DuckConnections(AZ1000,-1)', 'DuckConnections(AZ1000,2)')) {
        if (((Help-Text $invalid) | ConvertFrom-Json).state -ne 'failed') { throw "Invalid connection discovery input succeeded: $invalid" }
    }
    $endPage = (Help-Text 'DuckConnections(AZ1000,1)') | ConvertFrom-Json
    if ($endPage.state -ne 'succeeded' -or $endPage.connections.Count -ne 0 -or $null -ne $endPage.nextOffset) { throw 'Connection discovery pagination did not terminate.' }
    $afterDiscovery = Stats
    if ($afterDiscovery.session -ne $beforeDiscovery.session -or $afterDiscovery.statistics.Executions -ne $beforeDiscovery.statistics.Executions -or [string]$sheet.UsedRange.Address() -ne $beforeDiscoveryRange -or $sheet.Range('AZ1000').Value2) { throw 'Connection discovery changed the session, executed a query or wrote cells.' }
    $session.Workbook.Activate()
    Pass 'anchored discovery isolates workbooks, updates after configuration and performs no computation or cell edits'
    $beforeResources = (Stats).session
    $excel.Run('DuckSetResources', '{"MemoryMiB":1536,"SpillGiB":6,"Threads":1}', $session.Workbook.Name) | Out-Null
    [void](Wait-Status { param($s) $s.session -ne $beforeResources -and $s.resources.MemoryMiB -eq 1536 -and $s.connections -eq 1 })
    Wait-Value 'D1' 20
    if ($otherSheet.Range('A1').Value2 -ne $otherHandle) { throw 'Resource changes restarted another workbook.' }
    Pass 'resource changes preserve connections and restart only the selected workbook'
    $metadata = $session.Workbook.CustomXMLParts.SelectByNamespace('urn:xlduck:workbook-settings:v1')
    if ($metadata.Count -ne 1) { throw 'Connection definitions were not written into workbook metadata.' }
    $metadataXml = [xml]$metadata.Item(1).XML
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($metadata); $metadata = $null
    $definitions = $metadataXml.DocumentElement.FirstChild.InnerText | ConvertFrom-Json
    $connectionIdentity = $definitions.Connections[0].ConnectionId
    if (-not $connectionIdentity -or $definitions.Connections[0].Scope -ne 's3://fixture/' -or $definitions.Connections[0].PSObject.Properties.Name -contains 'CredentialId') { throw 'Portable metadata lost connection identity or embedded a vault reference.' }
    if ($session.Workbook.Saved) { throw 'Applying connection settings did not mark the workbook for saving.' }
    Pass 'connection definitions and resource budgets live in workbook metadata without vault references'
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
    # Model a workbook last saved with an older add-in. Only alter cached values in our closed fixture.
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($renamed, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry('xl/worksheets/sheet1.xml')
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $namespaces = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $namespaces.AddNamespace('s', $xml.DocumentElement.NamespaceURI)
        foreach ($address in @('AY2','AY3')) {
            $cell = $xml.SelectSingleNode('//s:c[@r="' + $address + '"]', $namespaces)
            if (-not $cell -or -not $cell.SelectSingleNode('s:f', $namespaces)) { throw 'Missing diagnostic formula fixture.' }
            $cell.SetAttribute('t','str')
            $value = $cell.SelectSingleNode('s:v', $namespaces)
            if (-not $value) { $value = $xml.CreateElement('v', $xml.DocumentElement.NamespaceURI); [void]$cell.AppendChild($value) }
            $value.InnerText = 'cached guidance from an older XlDuck build'
        }
        $entry.Delete(); $entry = $archive.CreateEntry('xl/worksheets/sheet1.xml')
        $writer = [IO.StreamWriter]::new($entry.Open())
        try { $writer.Write($xml.OuterXml) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    $session.Workbook = Invoke-XlDuckExcel { $excel.Workbooks.Open($renamed) }
    if ($null -eq $session.Workbook) {
        # Excel can open successfully without returning a COM result. Resolve only our exact fixture in our owned instance.
        foreach ($candidate in $excel.Workbooks) {
            if ($candidate.FullName -eq $renamed) { $session.Workbook = $candidate; break }
        }
    }
    if ($null -eq $session.Workbook) { throw 'Excel did not reopen the saved fixture.' }
    $sheet = $session.Workbook.Worksheets.Item(1)
    [void](Wait-Status { param($s) $s.ready -and $s.connections -eq 1 -and $s.session -ne $savedSession })
    Wait-Value 'F15' 100; Wait-Value 'D1' 20
    if ((Stats).resources.MemoryMiB -ne 1536) { throw 'Reopening lost the saved resource budget.' }
    Pass 'opening an ordinary saved workbook starts a fresh session and rereads its sources'
    Wait-Value 'AY3' (Help-Text 'DuckVersion()')
    Wait-Value 'AY2' (Help-Text 'DuckHelp()')
    Pass 'reopening refreshes help and version cells cached by an older add-in'
    $copy = Join-Path $caseDirectory 'unfamiliar copy.xlsx'
    $originalSession = (Stats).session
    [IO.File]::Copy($renamed, $copy)
    $copyBook = $excel.Workbooks.Open($copy); $copySheet = $copyBook.Worksheets.Item(1)
    try {
        [void](Wait-XlDuckCell $copySheet 'F15' -Until { param($v) $v -eq 100 })
        [void](Wait-XlDuckCell $copySheet 'Z1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).session -ne $originalSession -and ($v | ConvertFrom-Json).connections -eq 1 })
        $copyConnections = @($copySheet.Evaluate('DuckConnections(AZ1000)'))[0] | ConvertFrom-Json
        if ($copyConnections.connections[0].scope -ne 's3://fixture/') { throw 'A copied workbook lost its connections.' }
        $copyMetadata = [xml]$copyBook.CustomXMLParts.SelectByNamespace('urn:xlduck:workbook-settings:v1').Item(1).XML
        if (($copyMetadata.DocumentElement.FirstChild.InnerText | ConvertFrom-Json).Connections[0].ConnectionId -ne $connectionIdentity) { throw 'A copied workbook changed connection identity.' }
    } finally { $copyBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copySheet); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($copyBook) }
    Pass 'an externally copied workbook keeps its connection identities and computes in its own session'
    # Seed an old local record for a workbook created by this harness. The fake
    # vault reference tests association only; no credential is created or read.
    $legacyPath = Join-Path $caseDirectory 'legacy connections.xlsx'; $savedPaths += $legacyPath
    $migrationBook = $excel.Workbooks.Add(); $migrationBook.SaveAs($legacyPath, 51)
    $migrationBook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook); $migrationBook = $null
    [void](Stats)
    $fakeVaultId = [guid]::NewGuid().ToString('N')
    $legacySettings = @{ S3 = @(@{ Scope = 's3://migration/'; Credentials = 4; Endpoint = 'original.example.test'; CredentialId = $fakeVaultId }); Resources = @{ MemoryMiB = 2048; SpillGiB = 8; Threads = 3 } }
    $legacyCanonical = [IO.Path]::GetFullPath($legacyPath).ToUpperInvariant()
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $legacyHash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($legacyCanonical))).Replace('-', '') } finally { $sha.Dispose() }
    $legacyDirectory = Join-Path $env:LOCALAPPDATA 'XlDuck/connections'
    [void][IO.Directory]::CreateDirectory($legacyDirectory)
    [IO.File]::WriteAllText((Join-Path $legacyDirectory ($legacyHash + '.json')), (@{ Version = 1; Path = $legacyCanonical; Settings = $legacySettings } | ConvertTo-Json -Depth 8))
    $migrationBook = $excel.Workbooks.Open($legacyPath); $migrationSheet = $migrationBook.Worksheets.Item(1)
    $migrationSheet.Range('A1').Formula2 = '=DuckConnections()'
    $migrationResult = (Wait-XlDuckCell $migrationSheet 'A1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).state -eq 'succeeded' }) | ConvertFrom-Json
    if ($migrationResult.total -ne 1 -or $migrationResult.connections[0].requiresCredentialSetup) { throw 'Legacy settings did not acquire a local credential association.' }
    $migrationXml = [xml]$migrationBook.CustomXMLParts.SelectByNamespace('urn:xlduck:workbook-settings:v1').Item(1).XML
    $migrationDefinition = $migrationXml.DocumentElement.FirstChild.InnerText | ConvertFrom-Json
    $migrationId = $migrationDefinition.Connections[0].ConnectionId
    $bindingDirectory = Join-Path $env:LOCALAPPDATA 'XlDuck/credential-bindings'
    $bindingRecords = @(Get-ChildItem -LiteralPath $bindingDirectory -Filter '*.json' | Where-Object { ((Get-Content -LiteralPath $_.FullName -Raw) | ConvertFrom-Json).ConnectionId -eq $migrationId } | Select-Object -ExpandProperty FullName)
    if ($bindingRecords.Count -ne 1 -or $migrationXml.OuterXml.Contains($fakeVaultId) -or $migrationDefinition.Resources.MemoryMiB -ne 2048) { throw 'Migration lost budgets, exposed a vault reference or failed to persist the local association.' }
    $migrationBook.Save(); $migrationBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationSheet); $migrationSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook); $migrationBook = $null
    Pass 'legacy settings migrate into workbook metadata while the vault reference stays in the local association'
    $migratedCopy = Join-Path $caseDirectory 'migrated copy.xlsx'; [IO.File]::Copy($legacyPath, $migratedCopy)
    $migrationBook = $excel.Workbooks.Open($migratedCopy); $migrationSheet = $migrationBook.Worksheets.Item(1)
    $migrationResult = (Wait-XlDuckCell $migrationSheet 'A1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).state -eq 'succeeded' }) | ConvertFrom-Json
    if ($migrationResult.connections[0].requiresCredentialSetup) { throw 'Copying a workbook lost its credential association on this computer.' }
    # Change only the workbook's portable endpoint, bypassing the host dialog.
    $migrationDefinition.Connections[0].Endpoint = 'replacement.example.test'
    $migrationXml.DocumentElement.FirstChild.InnerText = $migrationDefinition | ConvertTo-Json -Depth 8 -Compress
    $migrationBook.CustomXMLParts.SelectByNamespace('urn:xlduck:workbook-settings:v1').Item(1).Delete()
    [void]$migrationBook.CustomXMLParts.Add($migrationXml.OuterXml)
    $migrationBook.Save(); $migrationBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationSheet); $migrationSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook); $migrationBook = $null
    [void](Stats)
    $migrationBook = $excel.Workbooks.Open($migratedCopy); $migrationSheet = $migrationBook.Worksheets.Item(1)
    $migrationResult = (Wait-XlDuckCell $migrationSheet 'A1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).state -eq 'succeeded' -and ($v | ConvertFrom-Json).connections[0].endpoint -eq 'replacement.example.test' }) | ConvertFrom-Json
    if (-not $migrationResult.connections[0].requiresCredentialSetup) { throw 'An endpoint edited in the workbook reused the original endpoint credential.' }
    $migrationDefinition.Version = 999
    $migrationXml.DocumentElement.FirstChild.InnerText = $migrationDefinition | ConvertTo-Json -Depth 8 -Compress
    $migrationBook.CustomXMLParts.SelectByNamespace('urn:xlduck:workbook-settings:v1').Item(1).Delete()
    [void]$migrationBook.CustomXMLParts.Add($migrationXml.OuterXml)
    $migrationBook.Save()
    $migrationBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationSheet); $migrationSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook); $migrationBook = $null
    Pass 'a copied workbook reuses local credentials only while its destination remains unchanged'
    [void](Stats)
    $migrationBook = $excel.Workbooks.Open($migratedCopy); $migrationSheet = $migrationBook.Worksheets.Item(1)
    [void](Wait-XlDuckCell $migrationSheet 'A1' -Until { param($v) $v -is [string] -and $v.StartsWith('{') -and ($v | ConvertFrom-Json).state -eq 'failed' })
    $migrationSheet.Range('A3').Formula2 = '=DuckQueryOutScalar("SELECT 1")'
    [void](Wait-XlDuckCell $migrationSheet 'A3' -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') -and $v.Contains('metadata') })
    if ((@($migrationSheet.Evaluate('DuckCellStatus(A3)'))[0] | ConvertFrom-Json).state -ne 'failed') { throw 'Invalid settings were reported as pending calculation.' }
    if ((@($migrationSheet.Evaluate('DuckExplore(AZ1000,"{""op"":""query"",""id"":""invalid-config"",""sql"":""SELECT 1""}")'))[0] | ConvertFrom-Json).state -ne 'failed') { throw 'Invalid settings left agent exploration initializing indefinitely.' }
    try { $excel.Run('DuckRestartWorkbook', $migrationBook.Name) | Out-Null } catch { }
    if ((@($migrationSheet.Evaluate('DuckConnections(AZ1000)'))[0] | ConvertFrom-Json).state -ne 'failed') { throw 'Restart silently replaced invalid metadata with defaults.' }
    $excel.Run('DuckSetConnections', '{"S3":[{"Scope":"s3://repaired/","Credentials":0}]}', $migrationBook.Name) | Out-Null
    [void](Wait-XlDuckCell $migrationSheet 'A3' -Until { param($v) $v -eq 1 })
    if ((@($migrationSheet.Evaluate('DuckConnections(AZ1000)'))[0] | ConvertFrom-Json).connections[0].scope -ne 's3://repaired/') { throw 'Explicitly applying valid settings did not repair metadata.' }
    $migrationBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationSheet); $migrationSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook); $migrationBook = $null
    Pass 'unsupported workbook metadata reports a configuration failure instead of indefinite initialization or default settings'
    $audit = $excel.Workbooks.Add(); $auditSheet = $audit.Worksheets.Item(1)
    try {
        $auditSheet.Range('Z1').Formula2 = '=DuckStatus()'
        $auditSheet.Range('A50').Formula2 = '=DuckQuery("SELECT 7654321")'
        [void](Wait-XlDuckCell $auditSheet 'A50' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$' })
        $auditSheet.Rows.Item(45).Insert() | Out-Null
        [void](Wait-XlDuckCell $auditSheet 'A51' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$' })
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
    if ($metadata) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($metadata); $metadata = $null }
    if ($migrationBook) { Invoke-XlDuckExcel { $migrationBook.Close($false) }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationBook) }
    if ($migrationSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($migrationSheet) }
    foreach ($bindingRecord in $bindingRecords) { Remove-Item -LiteralPath $bindingRecord }
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
