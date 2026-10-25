# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([switch]$NoBuild, [int]$InspectionSeconds = 0, [string]$AddInPath, [string]$RuntimeRoot, [string]$FrameworkProbe)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
if (-not $AddInPath) {
    if (-not $NoBuild) { Build-XlDuckAddIn }
    $AddInPath = Join-Path $PSScriptRoot '../XlDuck/bin/Debug/net8.0-windows/XlDuck-AddIn64.xll'
}
$build = Copy-XlDuckDevelopmentBuild $AddInPath
$session = $null; $sheet = $null
try {
    if ($RuntimeRoot) {
        $session = New-XlDuckExcelSession -AddInPath (Copy-XlDuckDevelopmentBuild $FrameworkProbe) -Visible
        $session.Application.Run('XlDuckProbeRuntimeRoot', [IO.Path]::GetFullPath($RuntimeRoot)) | Out-Null
        if (-not $session.Application.RegisterXLL($build)) { throw 'Production add-in could not load.' }
    } else { $session = New-XlDuckExcelSession -AddInPath $build -Visible }
    $excel = $session.Application; $sheet = $session.Workbook.Worksheets.Item(1)
    $session.Workbook.Windows.Item(1).Caption = 'XlDuck owned preview test'
    $sheet.Range('A1').Value2 = 'XlDuck preview integration'
    $sheet.Range('A3').Formula2 = '=DuckQuery("SELECT i::DOUBLE x, sin(i/10.0) y FROM range(100) r(i)")'
    $sheet.Range('A4').Formula2 = '=DuckPlot(A3,"line","x","x","y","y","title","Research preview")'
    [void](Wait-XlDuckCell $sheet 'A4' -Until { param($v) $v -is [string] -and $v -cmatch '^[TFP][1-9][0-9]*~[A-Za-z0-9_-]{11}$' })
    $sheet.Range('A4').Select() | Out-Null
    $excel.Run('DuckTogglePreview') | Out-Null
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $log = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'XlDuck') -Filter "xlduck-*-$($session.ProcessId).log" | Select-Object -First 1
        $text = if ($log) { Get-Content -LiteralPath $log.FullName -Raw } else { '' }
        if ($text -match 'NavigationCompleted, success=True' -and $text -match 'SetState: plot' -and $text -match 'DPI=\d+, ZoomFactor=') { break }
        if ($watch.Elapsed.TotalSeconds -gt 25) { throw 'The native preview did not become ready with a plot.' }
        Start-Sleep -Milliseconds 100
    } while ($true)
    Write-Host 'PASS Excel task pane starts its local WebView and receives the selected plot.'
    $scaling = [regex]::Match($text, 'DPI=(\d+), ZoomFactor=([0-9.,]+)')
    $zoom = [double]::Parse($scaling.Groups[2].Value.Replace(',','.'), [Globalization.CultureInfo]::InvariantCulture)
    if ([Math]::Abs($zoom - 1) -gt 0.0001) { throw "The preview reduced its page zoom at DPI $($scaling.Groups[1].Value): $zoom" }
    Write-Host "PASS native preview keeps 100% page zoom at DPI $($scaling.Groups[1].Value)."
    $before = (Get-Content -LiteralPath $log.FullName -Raw).Length
    $sheet.Range('A3').Formula2 = '=DuckQuery("SELECT i::DOUBLE x, cos(i/10.0) y FROM range(100) r(i)")'
    [void](Wait-XlDuckCell $sheet 'A4' -Until { param($v) $v -is [string] -and $v -cmatch '^[TFP][1-9][0-9]*~[A-Za-z0-9_-]{11}$' })
    $watch.Restart()
    do {
        $updates = (Get-Content -LiteralPath $log.FullName -Raw).Substring($before)
        if ($updates -match 'SetState: plot') { break }
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw 'The selected plot did not update after its dependency changed.' }
        Start-Sleep -Milliseconds 100
    } while ($true)
    if ($updates -match 'SetState: empty') { throw 'A transient dependency error cleared the previous chart.' }
    Write-Host 'PASS the selected chart updates without reselection and retains its previous state during dependency calculation.'
    if ($InspectionSeconds -gt 0) {
        Write-Host "Inspect owned Excel PID $($session.ProcessId); closing in $InspectionSeconds seconds."
        Start-Sleep -Seconds $InspectionSeconds
    }
    $excel.Run('DuckTogglePreview') | Out-Null
} finally {
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session) { Close-XlDuckExcelSession $session }
}
