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
    Write-TestResult "Returns handle format" ($handle -match "^duck://t/\d+$") "Got: $handle"

    # Test 2: range() function works
    Set-Formula "B1" '=DuckQuery("SELECT * FROM range(5)")'
    $handle2 = Get-CellValue "B1"
    Write-TestResult "range() query works" ($handle2 -match "^duck://t/\d+$") "Got: $handle2"
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
    Set-Formula "F1" '=DuckQueryOut("SELECT * FROM :src WHERE range > 2", "src", E1)'

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
    Set-Formula "B1" '=DuckQuery("SELECT COUNT(*) as cnt FROM :src", "src", A1)'
    Start-Sleep -Milliseconds 300
    Set-Formula "C1" "=DuckOut(B1)"

    $header = Get-CellValue "C1"
    $count = Get-CellValue "C2"

    Write-TestResult "Single param - header" ($header -eq "cnt") "Got: $header"
    Write-TestResult "Single param - count" ($count -eq "5") "Got: $count"

    # Test 2: Filtered query
    Set-Formula "D1" '=DuckQuery("SELECT * FROM :data WHERE range > 2", "data", A1)'
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

    Set-Formula "B1" '=DuckQuery("SELECT * FROM :src WHERE range >= 5", "src", A1)'
    Start-Sleep -Milliseconds 300

    Set-Formula "C1" '=DuckQuery("SELECT SUM(range) as total FROM :filtered", "filtered", B1)'
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
}

function Test-ErrorHandling {
    Write-Host "`nTest Suite: Error Handling" -ForegroundColor Cyan
    Clear-TestRange

    # Test 1: Invalid SQL
    Set-Formula "A1" '=DuckQuery("SELECT * FROM nonexistent_table")'
    $result = Get-CellValue "A1"
    Write-TestResult "Invalid table error" ($result -match "#ERROR") "Got: $result"

    # Test 2: Invalid handle
    Set-Formula "B1" '=DuckOut("duck://t/99999")'
    $result2 = Get-CellValue "B1"
    Write-TestResult "Invalid handle error" ($result2 -match "#ERROR") "Got: $result2"
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
Test-DuckExecute

# Summary
Write-Host "`n========================" -ForegroundColor White
Write-Host "Results: $script:TestsPassed passed, $script:TestsFailed failed" -ForegroundColor $(if ($script:TestsFailed -eq 0) { "Green" } else { "Red" })

if ($script:TestsFailed -gt 0) {
    exit 1
}
