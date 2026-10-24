# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0

# Development and integration sessions never attach through the ROT.
function Invoke-XlDuckExcel {
    param([Parameter(Mandatory)][scriptblock]$Action, [int]$TimeoutSeconds = 20)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        try { return (& $Action) }
        catch [Runtime.InteropServices.COMException] {
            # Excel rejects automation while processing a calculation or UI message.
            if ($_.Exception.HResult -notin @(-2147418111, -2147417846, -2146777998) -or
                $watch.Elapsed.TotalSeconds -ge $TimeoutSeconds) { throw }
            Start-Sleep -Milliseconds 25
        }
    }
}

function Get-XlDuckExcelProcessId {
    param([Parameter(Mandatory)]$Application)
    if (-not ('XlDuck.Testing.WindowProcess' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
namespace XlDuck.Testing {
    public static class WindowProcess {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    }
}
'@
    }
    [uint32]$excelProcessId = 0
    [void][XlDuck.Testing.WindowProcess]::GetWindowThreadProcessId(
        [IntPtr]$Application.Hwnd, [ref]$excelProcessId)
    if ($excelProcessId -eq 0) { throw 'Could not identify the Excel process.' }
    return [int]$excelProcessId
}

function Assert-XlDuckExcelSession {
    param([Parameter(Mandatory)]$Session)
    $ownedProcess = Get-Process -Id $Session.ProcessId -ErrorAction Stop
    if ($ownedProcess.ProcessName -ne 'EXCEL' -or
        $ownedProcess.StartTime.ToUniversalTime().Ticks -ne $Session.StartTimeTicks -or
        (Get-XlDuckExcelProcessId $Session.Application) -ne $Session.ProcessId) {
        throw 'Excel session ownership could not be verified. No workbook was changed.'
    }
}

function New-XlDuckExcelSession {
    param([string]$AddInPath, [switch]$Visible)
    $previousIds = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    $application = New-Object -ComObject Excel.Application
    $session = $null
    try {
        $excelProcessId = Get-XlDuckExcelProcessId $application
        if ($previousIds -contains $excelProcessId) {
            throw 'Excel reused an existing process. Refusing to modify or close it.'
        }
        $ownedProcess = Get-Process -Id $excelProcessId -ErrorAction Stop
        $session = [pscustomobject]@{
            Application = $application
            Workbook = $null
            ProcessId = $excelProcessId
            StartTimeTicks = $ownedProcess.StartTime.ToUniversalTime().Ticks
            WorkbookName = $null
        }
        # Startup workbooks belong to the user, even in a newly created process.
        if ($application.Workbooks.Count -ne 0) {
            throw 'Excel opened startup workbooks. Refusing to use this instance for tests.'
        }
        $application.Visible = [bool]$Visible
        $session.Workbook = $application.Workbooks.Add()
        $session.WorkbookName = $session.Workbook.Name
        Write-Host "Created owned Excel PID $excelProcessId, start ticks $($session.StartTimeTicks)." -ForegroundColor DarkGray
        if ($AddInPath) {
            $resolvedAddIn = (Resolve-Path -LiteralPath $AddInPath -ErrorAction Stop).Path
            if (-not $application.RegisterXLL($resolvedAddIn)) {
                throw "Excel could not load the add-in: $resolvedAddIn"
            }
        }
        return $session
    } catch {
        if ($session) { Close-XlDuckExcelSession $session }
        elseif ($application) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application) }
        throw
    }
}

function Close-XlDuckExcelSession {
    param([Parameter(Mandatory)]$Session)
    $quitRequested = $false
    try {
        Assert-XlDuckExcelSession $Session
        if ($Session.Workbook) {
            # Close the object the harness created, never ActiveWorkbook.
            Invoke-XlDuckExcel { $Session.Workbook.Close($false) }
        }
        if ($Session.Application.Workbooks.Count -eq 0) {
            Invoke-XlDuckExcel { $Session.Application.Quit() }
            $quitRequested = $true
        } else {
            Write-Warning "Leaving Excel PID $($Session.ProcessId) open: it contains other workbooks."
        }
    } finally {
        if ($Session.Workbook) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Session.Workbook)
            $Session.Workbook = $null
        }
        if ($Session.Application) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Session.Application)
            $Session.Application = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    if ($quitRequested) {
        $remaining = Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue
        if ($remaining -and $remaining.StartTime.ToUniversalTime().Ticks -eq $Session.StartTimeTicks -and
            -not $remaining.WaitForExit(5000)) {
            throw "Owned Excel PID $($Session.ProcessId) did not exit after Quit. It was not forcibly terminated."
        }
    }
}

function Wait-XlDuckCell {
    param(
        [Parameter(Mandatory)]$Sheet,
        [Parameter(Mandatory)][string]$Address,
        [scriptblock]$Until = { param($value) $null -ne $value -and $value -ne 'Loading...' },
        [int]$TimeoutSeconds = 20
    )
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastValue = $null
    do {
        try {
            $lastValue = $Sheet.Range($Address).Value2
            if (& $Until $lastValue) { return $lastValue }
        } catch [Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -notin @(-2147418111, -2147417846, -2146777998)) { throw }
        }
        Start-Sleep -Milliseconds 25
    } while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds)
    throw "Timed out waiting for $Address after ${TimeoutSeconds}s. Last value: $lastValue"
}

function Build-XlDuckAddIn {
    param([string]$ProjectPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'XlDuck/XlDuck.csproj'))
    & dotnet build $ProjectPath --nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed: $ProjectPath" }
}

function Copy-XlDuckDevelopmentBuild {
    param([Parameter(Mandatory)][string]$AddInPath)
    $source = (Resolve-Path -LiteralPath $AddInPath -ErrorAction Stop).Path
    $repo = Split-Path $PSScriptRoot -Parent
    $destination = Join-Path $repo ('artifacts/development/' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $destination -Force)
    foreach ($item in Get-ChildItem -LiteralPath (Split-Path $source -Parent)) {
        if ($item.Name -ne 'publish') {
            Copy-Item -LiteralPath $item.FullName -Destination $destination -Recurse
        }
    }
    return Join-Path $destination (Split-Path $source -Leaf)
}
