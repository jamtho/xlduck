# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([ValidateSet('inproc', 'worker')][string]$Mode = 'worker')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../../scripts/ExcelSession.ps1')
$session = $null
$sheet = $null
$otherBook = $null
$otherSheet = $null
$observations = [ordered]@{ mode = $Mode; checks = @() }
function Set-Formula([string]$address, [string]$formula) {
    Invoke-XlDuckExcel { $sheet.Range($address).Formula2 = $formula }
}
function Query([string]$sql, [string]$inputReference = '""') {
    return '=ProtoQuery("{0}","{1}",{2})' -f $Mode, $sql, $inputReference
}
function Info { return ($excel.Run('ProtoSessionInfo', $session.WorkbookName, $Mode) | ConvertFrom-Json) }
function Wait-State([scriptblock]$condition, [string]$description, [int]$seconds = 5) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (& $condition) { return }
        Start-Sleep -Milliseconds 25
    } while ($watch.Elapsed.TotalSeconds -lt $seconds)
    throw "Timed out: $description"
}
function Wait-Handle([string]$address) {
    return Wait-XlDuckCell $sheet $address -TimeoutSeconds 8 -Until {
        param($value)
        if ($value -is [string] -and $value.StartsWith('#ERROR:')) { throw $value }
        $value -is [string] -and $value.StartsWith('proto://')
    }
}
function Pass([string]$name) { Write-Host "PASS $name"; $observations.checks += $name }
try {
    $session = New-XlDuckExcelSession -AddInPath (Join-Path $PSScriptRoot 'AddIn/bin/Debug/net8.0-windows/AddIn-AddIn64.xll')
    $excel = $session.Application
    $sheet = $session.Workbook.Worksheets.Item(1)
    $excel.Run('ProtoSetWorker', (Join-Path $PSScriptRoot 'Worker/bin/Debug/net8.0/Worker.exe')) | Out-Null
    $excel.Run('ProtoSetPublisher', 'targeted', $session.WorkbookName) | Out-Null
    Set-Formula 'A1' (Query 'SELECT 1 AS value')
    [void](Wait-Handle 'A1')
    Write-Host 'Warm-up completed.'
    $initial = Info
    $observations.excelWorkingSetMiB = [Math]::Round((Get-Process -Id $session.ProcessId).WorkingSet64 / 1MB, 1)
    if ($initial.WorkerProcessId) {
        $observations.workerWorkingSetMiB = [Math]::Round((Get-Process -Id $initial.WorkerProcessId).WorkingSet64 / 1MB, 1)
    }

    # Branches share a materialized root; each result is published to its own requesting cell.
    $watch = [Diagnostics.Stopwatch]::StartNew()
    Set-Formula 'A2' (Query 'SELECT value+1 AS value FROM __input' 'A1')
    Set-Formula 'A3' (Query 'SELECT value+2 AS value FROM __input' 'A1')
    Set-Formula 'D2' ('=ProtoScalar("{0}",A2)' -f $Mode)
    Set-Formula 'D3' ('=ProtoScalar("{0}",A3)' -f $Mode)
    Set-Formula 'F1' '=D2+D3'
    [void](Wait-XlDuckCell $sheet 'F1' -Until { param($v) $v -is [double] -and $v -eq 5 })
    $observations.branchMs = $watch.Elapsed.TotalMilliseconds
    Pass 'branching DAG'

    $slow = 'SELECT sum(sin(i::DOUBLE)) AS value FROM range(1000000000) t(i)'
    $before = (Info).Executions
    Set-Formula 'A8' (Query $slow)
    Wait-State { $info = Info; $info.Executions -gt $before -and $info.Pending -gt 0 } 'slow request starts'
    $watch.Restart()
    Set-Formula 'A8' (Query 'SELECT 77 AS value')
    [void](Wait-Handle 'A8')
    $observations.replacePendingMs = $watch.Elapsed.TotalMilliseconds
    Set-Formula 'D8' ('=ProtoScalar("{0}",A8)' -f $Mode)
    [void](Wait-XlDuckCell $sheet 'D8' -Until { param($v) $v -is [double] -and $v -eq 77 })
    Pass 'replace pending formula and recover'

    $before = (Info).Executions
    Set-Formula 'A8' (Query $slow)
    Wait-State { $info = Info; $info.Executions -gt $before -and $info.Pending -gt 0 } 'second slow request starts'
    $sheet.Range('A8').ClearContents()
    Set-Formula 'A9' (Query 'SELECT 88 AS value')
    [void](Wait-Handle 'A9')
    if ($null -ne $sheet.Range('A8').Value2) { throw 'A late completion overwrote a deleted formula.' }
    Pass 'delete pending formula and recover'

    $before = (Info).Executions
    $shared = Query 'SELECT sum(sin(i::DOUBLE)) AS value FROM range(20000000) t(i)'
    Set-Formula 'A5' $shared
    Set-Formula 'A6' $shared
    Wait-State { (Info).Pending -gt 0 } 'shared request starts'
    $sheet.Range('A5').ClearContents()
    [void](Wait-Handle 'A6')
    if ((Info).Executions -ne $before + 1) { throw 'Shared work executed more than once.' }
    Pass 'one subscriber can leave shared work'

    $before = (Info).Executions
    Set-Formula 'A8' (Query $slow)
    Wait-State { $info = Info; $info.Executions -gt $before -and $info.Pending -gt 0 } 'request before restart starts'
    $old = Info
    $watch.Restart()
    $excel.Run('ProtoRestart', $session.WorkbookName) | Out-Null
    Set-Formula 'A8' (Query 'SELECT 99 AS value')
    [void](Wait-Handle 'A8')
    if ((Info).Id -eq $old.Id) { throw 'Restart kept the previous session.' }
    if ($old.WorkerProcessId) {
        Wait-State { $null -eq (Get-Process -Id $old.WorkerProcessId -ErrorAction SilentlyContinue) } 'old worker exits'
    }
    $observations.restartPendingMs = $watch.Elapsed.TotalMilliseconds
    Pass 'restart during pending work'

    # A second workbook can close during work without losing the first workbook's results.
    $retained = $sheet.Range('A1').Value2
    $otherBook = $excel.Workbooks.Add()
    $otherName = $otherBook.Name
    $otherSheet = $otherBook.Worksheets.Item(1)
    $otherSheet.Range('A1').Formula2 = Query $slow
    Wait-State { $state = $excel.Run('ProtoSessionInfo', $otherName, $Mode) | ConvertFrom-Json; $state.Pending -gt 0 } 'other workbook starts'
    $otherState = $excel.Run('ProtoSessionInfo', $otherName, $Mode) | ConvertFrom-Json
    $otherBook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet); $otherSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook); $otherBook = $null
    Wait-State { $excel.Run('ProtoSessionInfo', $otherName, $Mode) -eq '' } 'closed workbook session removed'
    if ($otherState.WorkerProcessId) {
        Wait-State { $null -eq (Get-Process -Id $otherState.WorkerProcessId -ErrorAction SilentlyContinue) } 'closed workbook worker exits'
    }
    $sheet.Range('A1').Calculate()
    if ($sheet.Range('A1').Value2 -ne $retained) { throw 'Closing another workbook changed a retained result.' }
    Pass 'close pending workbook without changing another'

    # Deliberately change only this owned Excel application's mode, then restore it.
    $excel.Calculation = -4135
    $sheet.Range('B20').Value2 = 10
    Set-Formula 'A20' ('=ProtoQuery("{0}","SELECT $value AS value","",B20)' -f $Mode)
    Set-Formula 'D20' ('=ProtoScalar("{0}",A20)' -f $Mode)
    $manualPasses = 0
    do {
        $manualPasses++
        $excel.Calculate()
        Wait-State { (Info).Pending -eq 0 } 'manual calculation work completes'
        # RTD notifications are still delivered in manual mode. Observe for one normal throttle period.
        $interval = [Diagnostics.Stopwatch]::StartNew()
        do { $value = $sheet.Range('D20').Value2; Start-Sleep -Milliseconds 25 }
        while ($value -ne 10 -and $interval.ElapsedMilliseconds -lt ($excel.RTD.ThrottleInterval + 250))
    } while ($value -ne 10 -and $manualPasses -lt 4)
    $observations.manualCalculationPasses = $manualPasses
    $observations.manualLeafValue = $value
    if ($excel.Calculation -ne -4135) { throw 'The add-in changed manual calculation mode.' }
    if ($value -ne 10) { throw 'Manual calculations did not converge.' }
    $excel.Calculation = -4105
    Pass 'manual calculation converges without changing application mode'

    if ($Mode -eq 'worker') {
        $before = (Info).Executions
        Set-Formula 'A25' (Query $slow)
        Wait-State { $info = Info; $info.Executions -gt $before -and $info.Pending -gt 0 } 'request before worker crash starts'
        $excel.Run('ProtoCrashWorker', $session.WorkbookName) | Out-Null
        [void](Wait-XlDuckCell $sheet 'A25' -TimeoutSeconds 5 -Until { param($v) $v -is [string] -and $v.StartsWith('#ERROR:') })
        Assert-XlDuckExcelSession $session
        $sheet.Range('A25').ClearContents()
        $excel.Run('ProtoRestart', $session.WorkbookName) | Out-Null
        [void](Wait-Handle 'A1')
        Pass 'worker crash is visible and restart recovers Excel'
    }
    $artifactDir = Join-Path $PSScriptRoot '../../artifacts/execution'
    [void](New-Item -ItemType Directory -Path $artifactDir -Force)
    $observations | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $artifactDir "$Mode-lifecycle.json") -Encoding UTF8
    $observations | ConvertTo-Json -Depth 6 | Write-Output
} catch {
    if ($sheet) {
        foreach ($address in @('A1', 'A2', 'A3', 'D2', 'D3', 'F1', 'A8', 'D20')) {
            try { Write-Host "$address : $($sheet.Range($address).Value2) [$($sheet.Range($address).Formula2)]" } catch { }
        }
    }
    throw
} finally {
    if ($session -and $session.Application) {
        try { $session.Application.Run('ProtoShutdown') | Out-Null } catch { Write-Warning $_ }
    }
    if ($otherBook) { $otherBook.Close($false) }
    if ($otherSheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherSheet) }
    if ($otherBook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($otherBook) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null }
    $excel = $null
    if ($session) { Close-XlDuckExcelSession $session }
}
