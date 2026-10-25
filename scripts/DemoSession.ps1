# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
. (Join-Path $PSScriptRoot 'ExcelSession.ps1')

function New-XlDuckDemoSession {
    param([Parameter(Mandatory)][string]$WorkbookPath, [string]$AddInPath, [switch]$CopyWorkbook)
    $repo = Split-Path $PSScriptRoot -Parent
    $workbook = (Resolve-Path -LiteralPath $WorkbookPath -ErrorAction Stop).Path
    if (-not $AddInPath) {
        $candidate = Get-ChildItem -LiteralPath (Join-Path $repo 'artifacts/packages') -Directory |
            Where-Object { (Test-Path -LiteralPath (Join-Path $_.FullName 'manifest.json')) -and (Test-Path -LiteralPath ($_.FullName + '.zip')) } |
            Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
        if (-not $candidate) { throw 'Build an XlDuck package or supply -AddInPath.' }
        & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $candidate.FullName
        $AddInPath = Join-Path $candidate.FullName 'XlDuck-AddIn64.xll'
    }
    if ($CopyWorkbook) {
        $directory = Join-Path $repo ('artifacts/demo-sessions/' + [guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($directory)
        $copy = Join-Path $directory ([IO.Path]::GetFileName($workbook))
        Copy-Item -LiteralPath $workbook -Destination $copy
        $workbook = $copy
    }
    $build = Copy-XlDuckDevelopmentBuild $AddInPath
    $runtime = Join-Path $repo 'artifacts/dotnet10/runtime'
    $probe = Join-Path $repo 'artifacts/runtime-framework/bin/XlDuck.RuntimeProbe/debug/XlDuck.RuntimeProbeFramework64.xll'
    $session = $null
    try {
        # The optional development bootstrap selects the portable runtime only in
        # our new process; it changes no machine settings or existing Excel process.
        if ((Test-Path -LiteralPath (Join-Path $runtime 'dotnet.exe')) -and (Test-Path -LiteralPath $probe)) {
            $session = New-XlDuckExcelSession -AddInPath (Copy-XlDuckDevelopmentBuild $probe)
            $session.Application.Run('XlDuckProbeRuntimeRoot', $runtime) | Out-Null
            if (-not $session.Application.RegisterXLL($build)) { throw 'Could not load XlDuck.' }
        } else { $session = New-XlDuckExcelSession -AddInPath $build }
        Assert-XlDuckExcelSession $session
        $blank = $session.Workbook
        $opened = Invoke-XlDuckExcel { $session.Application.Workbooks.Open($workbook) }
        $session.Workbook = $opened
        $session.WorkbookName = $opened.Name
        try { Invoke-XlDuckExcel { $blank.Close($false) } }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($blank) }
        if ($session.Application.Calculation -ne -4105) { throw 'This demo uses automatic calculation. Set that mode in a separate development Excel instance before trying again.' }
        return $session
    } catch {
        if ($session) { Close-XlDuckExcelSession $session }
        throw
    }
}
