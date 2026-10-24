# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$AddInPath, [Parameter(Mandatory)][string]$FrameworkProbe,
    [Parameter(Mandatory)][string]$CoreProbe, [Parameter(Mandatory)][string]$ProbeFunction,
    [string]$RuntimeRoot, [int]$ExpectedMajor, [string[]]$Orders = @('xlduck-first','companion-first'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
foreach ($order in $Orders) {
    $session = $null; $sheet = $null
    try {
        $session = New-XlDuckExcelSession -AddInPath (Copy-XlDuckDevelopmentBuild $FrameworkProbe)
        $excel = $session.Application; $sheet = $session.Workbook.Worksheets.Item(1)
        if ($RuntimeRoot) { $excel.Run('XlDuckProbeRuntimeRoot', [IO.Path]::GetFullPath($RuntimeRoot)) | Out-Null }
        $paths = if ($order -eq 'xlduck-first') { @($AddInPath,$CoreProbe) } else { @($CoreProbe,$AddInPath) }
        foreach ($path in $paths) {
            if (-not $excel.RegisterXLL((Copy-XlDuckDevelopmentBuild $path))) { throw "Add-in load failed: $path" }
        }
        $excel.Run('DuckEnableWorkbook', $session.Workbook.Name) | Out-Null
        $sheet.Range('A1').Formula2 = '=DuckQueryOutScalar("SELECT 42")'
        $sheet.Range('B1').Formula2 = '=' + $ProbeFunction + '()'
        $sheet.Range('C1').Formula2 = '=XlDuckProbeFramework()'
        $sheet.Range('D1').Formula2 = '=DuckStatus()'
        [void](Wait-XlDuckCell $sheet 'A1' -Until { param($v) $v -ceq 42d })
        Write-Host "Loaded production host: $($sheet.Range('D1').Value2)"
        $probeRuntime = Wait-XlDuckCell $sheet 'B1' -Until { param($v) $v -is [string] -and $v -match '^\d+\.' }
        $status = $sheet.Range('D1').Value2 | ConvertFrom-Json
        if ($ExpectedMajor -and ([version]$status.runtime).Major -ne $ExpectedMajor) { throw "Unexpected XlDuck runtime $($status.runtime)" }
        if (-not $sheet.Range('C1').Value2.StartsWith('4.')) { throw '.NET Framework companion did not calculate.' }
        Write-Host "PASS $order with $ProbeFunction`: XlDuck runtime $($status.runtime), companion $probeRuntime, .NET Framework companion active."
    } finally {
        if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
        if ($session) { Close-XlDuckExcelSession $session }
    }
}
