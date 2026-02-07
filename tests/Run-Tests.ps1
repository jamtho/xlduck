# Copyright (c) 2026 James Thompson
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

# XlDuck Integration Tests
# Requires: Excel running with XlDuck add-in loaded and a workbook open

param(
    [switch]$Verbose
)

$script:TestsPassed = 0
$script:TestsFailed = 0
$script:Excel = $null
$script:Sheet = $null

function Write-TestResult {
    param([string]$Name, [bool]$Passed, [string]$Details = "")
    if ($Passed) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        $script:TestsPassed++
    } else {
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Details) { Write-Host "         $Details" -ForegroundColor Yellow }
        $script:TestsFailed++
    }
}

function Initialize-Excel {
    try {
        $script:Excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $script:Sheet = $script:Excel.ActiveWorkbook.ActiveSheet
        return $true
    } catch {
        Write-Host "ERROR: Could not connect to Excel. Make sure Excel is running with the add-in loaded." -ForegroundColor Red
        Write-Host "       Launch: Start-Process 'XlDuck\bin\Debug\net8.0-windows\XlDuck-AddIn64.xll'" -ForegroundColor Yellow
        return $false
    }
}

function Clear-TestRange {
    $script:Sheet.Range("A1:Z50").Clear() | Out-Null
    Start-Sleep -Milliseconds 200
}

function Set-Formula {
    param([string]$Cell, [string]$Formula)
    $script:Sheet.Range($Cell).Formula2 = $Formula
    Start-Sleep -Milliseconds 300
}

function Get-CellValue {
    param([string]$Cell)
    return $script:Sheet.Range($Cell).Text
}

function Get-CellFormula {
    param([string]$Cell)
    return $script:Sheet.Range($Cell).Formula
}

# ============================================
# Test Suites
# ============================================

function Test-Versions {
    Write-Host "`nTest Suite: Version Functions" -ForegroundColor Cyan
    Clear-TestRange

    Set-Formula "A1" "=DuckVersion()"
    $addinVersion = Get-CellValue "A1"
    Write-TestResult "DuckVersion returns add-in version" ($addinVersion -eq "0.1") "Got: $addinVersion"

    Set-Formula "B1" "=DuckLibraryVersion()"
    $libVersion = Get-CellValue "B1"
    Write-TestResult "DuckLibraryVersion returns DuckDB version" ($libVersion -match "^v\d+\.\d+\.\d+$") "Got: $libVersion"
}

function Test-DuckQueryBasic {
    Write-Host "`nTest Suite: DuckQuery Basic" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Simple query returns handle
    Set-Formula "A1" '=DuckQuery("SELECT 1 as num")'
    $handle = Get-CellValue "A1"
    Write-TestResult "Returns handle format" ($handle -match "^duck://table/\d+\|\d+x\d+$") "Got: $handle"

    # Test 2: range() function works
    Set-Formula "B1" '=DuckQuery("SELECT * FROM range(5)")'
    $handle2 = Get-CellValue "B1"
    Write-TestResult "range() query works" ($handle2 -match "^duck://table/\d+\|\d+x\d+$") "Got: $handle2"
}

function Test-DuckOut {
    Write-Host "`nTest Suite: DuckOut" -ForegroundColor Cyan
    Clear-TestRange

    # Setup: Create a handle
    Set-Formula "A1" '=DuckQuery("SELECT * FROM range(3)")'
    Start-Sleep -Milliseconds 300

    # Test 1: Spills array with header
    Set-Formula "B1" "=DuckOut(A1)"
    $header = Get-CellValue "B1"
    $val1 = Get-CellValue "B2"
    $val2 = Get-CellValue "B3"
    $val3 = Get-CellValue "B4"

    Write-TestResult "Header row present" ($header -eq "range") "Got: $header"
    Write-TestResult "First value correct" ($val1 -eq "0") "Got: $val1"
    Write-TestResult "Second value correct" ($val2 -eq "1") "Got: $val2"
    Write-TestResult "Third value correct" ($val3 -eq "2") "Got: $val3"

    # Test 2: Multi-column output
    Set-Formula "D1" '=DuckQuery("SELECT 1 as a, 2 as b, 3 as c")'
    Start-Sleep -Milliseconds 300
    Set-Formula "E1" "=DuckOut(D1)"

    $colA = Get-CellValue "E1"
    $colB = Get-CellValue "F1"
    $colC = Get-CellValue "G1"

    Write-TestResult "Multi-column headers" ($colA -eq "a" -and $colB -eq "b" -and $colC -eq "c") "Got: $colA, $colB, $colC"
}

function Test-DuckQueryOut {
    Write-Host "`nTest Suite: DuckQueryOut (combo function)" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Direct query to output
    Set-Formula "A1" '=DuckQueryOut("SELECT * FROM range(3)")'
    $header = Get-CellValue "A1"
    $val1 = Get-CellValue "A2"
    $val2 = Get-CellValue "A3"
    $val3 = Get-CellValue "A4"

    Write-TestResult "Header row present" ($header -eq "range") "Got: $header"
    Write-TestResult "First value correct" ($val1 -eq "0") "Got: $val1"
    Write-TestResult "Second value correct" ($val2 -eq "1") "Got: $val2"
    Write-TestResult "Third value correct" ($val3 -eq "2") "Got: $val3"

    # Test 2: With parameter binding
    Set-Formula "E1" '=DuckQuery("SELECT * FROM range(5)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "F1" '=DuckQueryOut("SELECT * FROM ? WHERE range > 2", E1)'

    $filtered1 = Get-CellValue "F2"
    $filtered2 = Get-CellValue "F3"

    Write-TestResult "Parameter binding works" ($filtered1 -eq "3" -and $filtered2 -eq "4") "Got: $filtered1, $filtered2"
}

function Test-ParameterBinding {
    Write-Host "`nTest Suite: Parameter Binding" -ForegroundColor Cyan
    Clear-TestRange

    # Setup: Create source data
    Set-Formula "A1" '=DuckQuery("SELECT * FROM range(5)")'
    Start-Sleep -Milliseconds 300

    # Test 1: Single parameter
    Set-Formula "B1" '=DuckQuery("SELECT COUNT(*) as cnt FROM ?", A1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "C1" "=DuckOut(B1)"

    $header = Get-CellValue "C1"
    $count = Get-CellValue "C2"

    Write-TestResult "Single param - header" ($header -eq "cnt") "Got: $header"
    Write-TestResult "Single param - count" ($count -eq "5") "Got: $count"

    # Test 2: Filtered query
    Set-Formula "D1" '=DuckQuery("SELECT * FROM ? WHERE range > 2", A1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "E1" "=DuckOut(D1)"

    $val1 = Get-CellValue "E2"
    $val2 = Get-CellValue "E3"

    Write-TestResult "Filter - first value" ($val1 -eq "3") "Got: $val1"
    Write-TestResult "Filter - second value" ($val2 -eq "4") "Got: $val2"
}

function Test-ChainedQueries {
    Write-Host "`nTest Suite: Chained Queries" -ForegroundColor Cyan
    Clear-TestRange

    # Setup: Create chain A1 -> B1 -> C1
    Set-Formula "A1" '=DuckQuery("SELECT * FROM range(10)")'
    Start-Sleep -Milliseconds 300

    Set-Formula "B1" '=DuckQuery("SELECT * FROM ? WHERE range >= 5", A1)'
    Start-Sleep -Milliseconds 300

    Set-Formula "C1" '=DuckQuery("SELECT SUM(range) as total FROM ?", B1)'
    Start-Sleep -Milliseconds 300

    Set-Formula "D1" "=DuckOut(C1)"

    $header = Get-CellValue "D1"
    $sum = Get-CellValue "D2"

    # Sum of 5+6+7+8+9 = 35
    Write-TestResult "Chain result header" ($header -eq "total") "Got: $header"
    Write-TestResult "Chain result value" ($sum -eq "35") "Got: $sum (expected 35 = 5+6+7+8+9)"
}

function Test-TypeConversions {
    Write-Host "`nTest Suite: Type Conversions" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: HUGEINT/SUM conversion
    Set-Formula "A1" '=DuckQuery("SELECT SUM(range) as total FROM range(100)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "B1" "=DuckOut(A1)"

    $sum = Get-CellValue "B2"
    # Sum of 0..99 = 4950
    Write-TestResult "HUGEINT SUM converts" ($sum -eq "4950") "Got: $sum"

    # Test 2: String values
    Set-Formula "C1" "=DuckQuery(""SELECT 'hello' as greeting"")"
    Start-Sleep -Milliseconds 300
    Set-Formula "D1" "=DuckOut(C1)"

    $str = Get-CellValue "D2"
    Write-TestResult "String values work" ($str -eq "hello") "Got: $str"

    # Test 3: NULL handling
    Set-Formula "E1" '=DuckQuery("SELECT NULL as empty")'
    Start-Sleep -Milliseconds 300
    Set-Formula "F1" "=DuckOut(E1)"

    $null_val = Get-CellValue "F2"
    Write-TestResult "NULL becomes empty" ($null_val -eq "") "Got: '$null_val'"

    # Test 4: TINYINT (sbyte) handling - must work through temp table creation
    Set-Formula "G1" '=DuckQuery("SELECT 1::TINYINT as val")'
    Start-Sleep -Milliseconds 300
    Set-Formula "H1" '=DuckQuery("SELECT val + 1 as result FROM ?", G1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "I1" "=DuckOut(H1)"

    $tinyint_val = Get-CellValue "I2"
    Write-TestResult "TINYINT (sbyte) works" ($tinyint_val -eq "2") "Got: $tinyint_val"

    # Test 5: COUNT(DISTINCT) through temp table
    Set-Formula "J1" '=DuckQuery("SELECT * FROM (VALUES (1), (2), (2), (3), (3), (3)) AS t(id)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "K1" '=DuckQuery("SELECT COUNT(DISTINCT id) as cnt FROM ?", J1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "L1" "=DuckOut(K1)"

    $cnt = Get-CellValue "L2"
    Write-TestResult "COUNT(DISTINCT) works" ($cnt -eq "3") "Got: $cnt"
}

function Test-ErrorHandling {
    Write-Host "`nTest Suite: Error Handling" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Invalid table - should be notfound category
    Set-Formula "A1" '=DuckQuery("SELECT * FROM nonexistent_table")'
    $result = Get-CellValue "A1"
    Write-TestResult "Invalid table error (notfound)" ($result -match "#duck://error/notfound\|") "Got: $result"

    # Test 2: Invalid handle - should be notfound category
    Set-Formula "B1" '=DuckOut("duck://table/99999")'
    $result2 = Get-CellValue "B1"
    Write-TestResult "Invalid handle error (notfound)" ($result2 -match "#duck://error/notfound\|") "Got: $result2"

    # Test 3: Syntax error
    Set-Formula "C1" '=DuckQuery("SELEC * FORM table")'
    $result3 = Get-CellValue "C1"
    Write-TestResult "Syntax error category" ($result3 -match "#duck://error/syntax\|") "Got: $result3"
}

function Test-DuckConfigReady {
    Write-Host "`nTest Suite: DuckConfigReady" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: DuckConfigReady returns OK
    Set-Formula "A1" '=DuckConfigReady()'
    $result = Get-CellValue "A1"
    Write-TestResult "DuckConfigReady returns OK" ($result -eq "OK") "Got: $result"

    # Test 2: DuckFragAfterConfig works alongside real parameters
    Set-Formula "B1" '=DuckQuery("SELECT * FROM range(3)")'
    Start-Sleep -Milliseconds 500
    Set-Formula "C1" '=DuckFragAfterConfig("SELECT * FROM ?", B1)'
    $result2 = Get-CellValue "C1"
    Write-TestResult "DuckFragAfterConfig with params returns handle" ($result2 -match "^duck://frag/\d+$") "Got: $result2"
}

function Test-DuckExecute {
    Write-Host "`nTest Suite: DuckExecute" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Create table
    Set-Formula "A1" '=DuckExecute("CREATE TABLE test_tbl (id INT, name VARCHAR)")'
    $result = Get-CellValue "A1"
    Write-TestResult "CREATE TABLE succeeds" ($result -match "OK") "Got: $result"

    # Test 2: Insert data
    Set-Formula "B1" "=DuckExecute(""INSERT INTO test_tbl VALUES (1, 'alice'), (2, 'bob')"")"
    $result2 = Get-CellValue "B1"
    Write-TestResult "INSERT succeeds" ($result2 -match "OK.*2 rows") "Got: $result2"

    # Test 3: Query the table
    Set-Formula "C1" '=DuckQuery("SELECT * FROM test_tbl ORDER BY id")'
    Start-Sleep -Milliseconds 300
    Set-Formula "D1" "=DuckOut(C1)"

    $id1 = Get-CellValue "D2"
    $name1 = Get-CellValue "E2"

    Write-TestResult "Query created table" ($id1 -eq "1" -and $name1 -eq "alice") "Got: id=$id1, name=$name1"

    # Cleanup
    Set-Formula "E1" '=DuckExecute("DROP TABLE test_tbl")'
}

function Test-DuckFragBasic {
    Write-Host "`nTest Suite: DuckFrag Basic" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Simple fragment returns handle
    Set-Formula "A1" '=DuckFrag("SELECT * FROM range(5)")'
    $handle = Get-CellValue "A1"
    Write-TestResult "Returns fragment handle format" ($handle -match "^duck://frag/\d+$") "Got: $handle"

    # Test 2: Fragment with invalid SQL returns error at creation time
    Set-Formula "B1" '=DuckFrag("SELECT * FROM nonexistent_table")'
    $result = Get-CellValue "B1"
    Write-TestResult "Invalid SQL detected at creation" ($result -match "#duck://error/") "Got: $result"
}

function Test-DuckFragInQuery {
    Write-Host "`nTest Suite: DuckFrag in DuckQuery" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Fragment used as table source
    Set-Formula "A1" '=DuckFrag("SELECT * FROM range(5)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "B1" '=DuckQuery("SELECT * FROM ? WHERE range > 2", A1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "C1" "=DuckOut(B1)"

    $val1 = Get-CellValue "C2"
    $val2 = Get-CellValue "C3"

    Write-TestResult "Fragment in query - first value" ($val1 -eq "3") "Got: $val1"
    Write-TestResult "Fragment in query - second value" ($val2 -eq "4") "Got: $val2"

    # Test 2: Fragment with aggregation
    Set-Formula "D1" '=DuckFrag("SELECT * FROM range(10)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "E1" '=DuckQuery("SELECT SUM(range) as total FROM ?", D1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "F1" "=DuckOut(E1)"

    $sum = Get-CellValue "F2"
    # Sum of 0..9 = 45
    Write-TestResult "Fragment aggregation works" ($sum -eq "45") "Got: $sum"
}

function Test-DuckFragChained {
    Write-Host "`nTest Suite: DuckFrag Chaining" -ForegroundColor Cyan
    Clear-TestRange

    # Chain: A1 (frag) -> B1 (frag) -> C1 (query)
    Set-Formula "A1" '=DuckFrag("SELECT * FROM range(10)")'
    Start-Sleep -Milliseconds 300

    Set-Formula "B1" '=DuckFrag("SELECT * FROM ? WHERE range >= 5", A1)'
    Start-Sleep -Milliseconds 300

    Set-Formula "C1" '=DuckQuery("SELECT SUM(range) as total FROM ?", B1)'
    Start-Sleep -Milliseconds 300

    Set-Formula "D1" "=DuckOut(C1)"

    $sum = Get-CellValue "D2"
    # Sum of 5+6+7+8+9 = 35
    Write-TestResult "Chained fragments - sum" ($sum -eq "35") "Got: $sum (expected 35)"
}

function Test-DuckFragWithTableHandle {
    Write-Host "`nTest Suite: DuckFrag with Table Handle" -ForegroundColor Cyan
    Clear-TestRange

    # Create a table handle first
    Set-Formula "A1" '=DuckQuery("SELECT * FROM range(5)")'
    Start-Sleep -Milliseconds 300

    # Fragment references the table handle
    Set-Formula "B1" '=DuckFrag("SELECT * FROM ? WHERE range > 2", A1)'
    Start-Sleep -Milliseconds 300

    # Query the fragment
    Set-Formula "C1" '=DuckQuery("SELECT COUNT(*) as cnt FROM ?", B1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "D1" "=DuckOut(C1)"

    $count = Get-CellValue "D2"
    Write-TestResult "Fragment with table handle" ($count -eq "2") "Got: $count (expected 2 rows: 3, 4)"
}

function Test-DuckOutWithFragment {
    Write-Host "`nTest Suite: DuckOut with Fragment Handle" -ForegroundColor Cyan
    Clear-TestRange

    # DuckOut can directly output a fragment
    Set-Formula "A1" '=DuckFrag("SELECT * FROM range(3)")'
    Start-Sleep -Milliseconds 300
    Set-Formula "B1" "=DuckOut(A1)"

    $header = Get-CellValue "B1"
    $val1 = Get-CellValue "B2"
    $val2 = Get-CellValue "B3"
    $val3 = Get-CellValue "B4"

    Write-TestResult "DuckOut fragment - header" ($header -eq "range") "Got: $header"
    Write-TestResult "DuckOut fragment - first value" ($val1 -eq "0") "Got: $val1"
    Write-TestResult "DuckOut fragment - second value" ($val2 -eq "1") "Got: $val2"
    Write-TestResult "DuckOut fragment - third value" ($val3 -eq "2") "Got: $val3"
}

function Test-Pivot {
    Write-Host "`nTest Suite: PIVOT" -ForegroundColor Cyan
    Clear-TestRange

    # Create source data
    Set-Formula "A1" "=DuckFrag(""SELECT * FROM (VALUES ('Q1','North',100), ('Q1','South',150), ('Q2','North',200), ('Q2','South',250)) AS sales(quarter, region, amount)"")"
    Start-Sleep -Milliseconds 300

    # Pivot the data
    Set-Formula "B1" '=DuckQueryOut("PIVOT ? ON region USING SUM(amount) ORDER BY quarter", A1)'

    $header1 = Get-CellValue "B1"
    $header2 = Get-CellValue "C1"
    $header3 = Get-CellValue "D1"

    Write-TestResult "PIVOT headers" ($header1 -eq "quarter" -and $header2 -eq "North" -and $header3 -eq "South") "Got: $header1, $header2, $header3"

    $q1North = Get-CellValue "C2"
    $q1South = Get-CellValue "D2"
    $q2North = Get-CellValue "C3"
    $q2South = Get-CellValue "D3"

    Write-TestResult "PIVOT values" ($q1North -eq "100" -and $q1South -eq "150" -and $q2North -eq "200" -and $q2South -eq "250") "Got: Q1=($q1North,$q1South), Q2=($q2North,$q2South)"
}

function Test-DuckCapture {
    Write-Host "`nTest Suite: DuckCapture" -ForegroundColor Cyan
    Clear-TestRange

    # Setup: Put data into a range (headers + data)
    $script:Sheet.Range("A1").Value2 = "name"
    $script:Sheet.Range("B1").Value2 = "age"
    $script:Sheet.Range("C1").Value2 = "city"
    $script:Sheet.Range("A2").Value2 = "alice"
    $script:Sheet.Range("B2").Value2 = 30
    $script:Sheet.Range("C2").Value2 = "NYC"
    $script:Sheet.Range("A3").Value2 = "bob"
    $script:Sheet.Range("B3").Value2 = 25
    $script:Sheet.Range("C3").Value2 = "LA"
    $script:Sheet.Range("A4").Value2 = "charlie"
    $script:Sheet.Range("B4").Value2 = 35
    $script:Sheet.Range("C4").Value2 = "Chicago"
    Start-Sleep -Milliseconds 200

    # Test 1: Basic capture returns handle
    Set-Formula "E1" '=DuckCapture(A1:C4)'
    $handle = Get-CellValue "E1"
    Write-TestResult "Returns table handle format" ($handle -match "^duck://table/\d+\|3x3$") "Got: $handle"

    # Test 2: DuckOut of captured data preserves headers and values
    Set-Formula "F1" '=DuckOut(E1)'
    Start-Sleep -Milliseconds 300
    $hdr1 = Get-CellValue "F1"
    $hdr2 = Get-CellValue "G1"
    $hdr3 = Get-CellValue "H1"
    $val1 = Get-CellValue "F2"
    $val2 = Get-CellValue "G2"
    $val3 = Get-CellValue "H2"

    Write-TestResult "Headers preserved" ($hdr1 -eq "name" -and $hdr2 -eq "age" -and $hdr3 -eq "city") "Got: $hdr1, $hdr2, $hdr3"
    Write-TestResult "First row values" ($val1 -eq "alice" -and $val2 -eq "30" -and $val3 -eq "NYC") "Got: $val1, $val2, $val3"

    # Test 3: SQL query on captured data
    Set-Formula "J1" '=DuckQueryOut("SELECT * FROM ? WHERE age > 28", E1)'
    Start-Sleep -Milliseconds 300
    $qName1 = Get-CellValue "J2"
    $qName2 = Get-CellValue "J3"

    Write-TestResult "SQL filter on captured data" ($qName1 -eq "alice" -and $qName2 -eq "charlie") "Got: $qName1, $qName2"

    # Test 4: Numeric type inference (SUM works)
    Set-Formula "M1" '=DuckQueryOut("SELECT SUM(age) as total FROM ?", E1)'
    Start-Sleep -Milliseconds 300
    $sum = Get-CellValue "M2"
    Write-TestResult "Numeric type inference (SUM)" ($sum -eq "90") "Got: $sum (expected 90 = 30+25+35)"

    # Test 5: Empty cells become NULL
    Clear-TestRange
    $script:Sheet.Range("A1").Value2 = "val"
    $script:Sheet.Range("A2").Value2 = 10
    # A3 intentionally left empty
    $script:Sheet.Range("A4").Value2 = 30
    Start-Sleep -Milliseconds 200

    Set-Formula "B1" '=DuckCapture(A1:A4)'
    Start-Sleep -Milliseconds 300
    Set-Formula "C1" '=DuckQueryOut("SELECT COUNT(val) as cnt FROM ?", B1)'
    Start-Sleep -Milliseconds 300
    $cnt = Get-CellValue "C2"
    Write-TestResult "Empty cells become NULL (COUNT skips)" ($cnt -eq "2") "Got: $cnt (expected 2, skipping NULL)"
}

function Test-ReadCSV {
    Write-Host "`nTest Suite: Read CSV Files" -ForegroundColor Cyan
    Clear-TestRange

    # Get the path to the test CSV (relative to repo root)
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $csvPath = Join-Path $repoRoot "tests\data\sample.csv"
    $csvPath = $csvPath -replace '\\', '/'  # DuckDB prefers forward slashes

    # Test 1: Read CSV with read_csv_auto
    Set-Formula "A1" "=DuckQueryOut(""SELECT * FROM read_csv_auto('$csvPath')"")"

    $header1 = Get-CellValue "A1"
    $header2 = Get-CellValue "B1"
    $header3 = Get-CellValue "C1"

    Write-TestResult "CSV headers detected" ($header1 -eq "id" -and $header2 -eq "name" -and $header3 -eq "value") "Got: $header1, $header2, $header3"

    $name1 = Get-CellValue "B2"
    $name2 = Get-CellValue "B3"
    $name3 = Get-CellValue "B4"

    Write-TestResult "CSV data rows" ($name1 -eq "alice" -and $name2 -eq "bob" -and $name3 -eq "charlie") "Got: $name1, $name2, $name3"

    # Test 2: Query CSV with filter
    Set-Formula "E1" "=DuckQueryOut(""SELECT name, value FROM read_csv_auto('$csvPath') WHERE value > 150"")"

    $filteredName = Get-CellValue "E2"
    $filteredValue = Get-CellValue "F2"

    Write-TestResult "CSV filtered query" ($filteredName -eq "bob" -and $filteredValue -eq "200") "Got: name=$filteredName, value=$filteredValue"

    # Test 3: String parameter binding (path as parameter)
    Clear-TestRange
    $script:Sheet.Range("A1").Value2 = $csvPath
    Set-Formula "B1" '=DuckFrag("SELECT * FROM read_csv_auto(?)", A1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "C1" '=DuckQueryOut("SELECT name FROM ? WHERE id = 1", B1)'

    $name = Get-CellValue "C2"
    Write-TestResult "String param in read_csv_auto" ($name -eq "alice") "Got: $name"
}

# ============================================
# Main
# ============================================

Write-Host "XlDuck Integration Tests" -ForegroundColor White
Write-Host "========================" -ForegroundColor White

if (-not (Initialize-Excel)) {
    exit 1
}

Write-Host "Connected to Excel: $($script:Excel.ActiveWorkbook.Name)" -ForegroundColor Gray

# Run all test suites
Test-Versions
Test-DuckQueryBasic
Test-DuckOut
Test-DuckQueryOut
Test-ParameterBinding
Test-ChainedQueries
Test-TypeConversions
Test-ErrorHandling
Test-DuckConfigReady
Test-DuckExecute
Test-DuckFragBasic
Test-DuckFragInQuery
Test-DuckFragChained
Test-DuckFragWithTableHandle
Test-DuckOutWithFragment
Test-Pivot
Test-DuckCapture
Test-ReadCSV

# Summary
Write-Host "`n========================" -ForegroundColor White
Write-Host "Results: $script:TestsPassed passed, $script:TestsFailed failed" -ForegroundColor $(if ($script:TestsFailed -eq 0) { "Green" } else { "Red" })

if ($script:TestsFailed -gt 0) {
    exit 1
}
