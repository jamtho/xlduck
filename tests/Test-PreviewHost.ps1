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
    [void](Wait-XlDuckCell $sheet 'A4' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$' })
    $sheet.Range('A4').Select() | Out-Null
    $excel.Run('DuckTogglePreview') | Out-Null
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $log = Get-ChildItem (Join-Path $env:LOCALAPPDATA 'XlDuck') -Filter "xlduck-*-$($session.ProcessId).log" | Select-Object -First 1
        $text = if ($log) { Get-Content -LiteralPath $log.FullName -Raw } else { '' }
        if ($text -match 'NavigationCompleted, success=True' -and $text -match 'SetState: plot' -and $text -match 'ClientWidth=\d+, Metrics=') { break }
        if ($watch.Elapsed.TotalSeconds -gt 25) { throw 'The native preview did not become ready with a plot.' }
        Start-Sleep -Milliseconds 100
    } while ($true)
    Write-Host 'PASS Excel task pane starts its local WebView and receives the selected plot.'
    $scaling = [regex]::Match($text, 'DPI=(\d+), ZoomFactor=([0-9.,]+)')
    $zoom = [double]::Parse($scaling.Groups[2].Value.Replace(',','.'), [Globalization.CultureInfo]::InvariantCulture)
    if ([Math]::Abs($zoom - 1) -gt 0.0001) { throw "The preview reduced its page zoom at DPI $($scaling.Groups[1].Value): $zoom" }
    Write-Host "PASS native preview keeps 100% page zoom at DPI $($scaling.Groups[1].Value)."
    $measurement = [regex]::Match($text, 'WindowDPI=(\d+), ClientWidth=(\d+), Metrics=(\{[^\r\n]+\})')
    $metrics = $measurement.Groups[3].Value | ConvertFrom-Json
    $expectedScale = [double]$measurement.Groups[1].Value / 96
    if ([Math]::Abs($metrics.dpr - $expectedScale) -gt 0.01 -or
        [Math]::Abs($metrics.width * $expectedScale - [double]$measurement.Groups[2].Value) -gt 3) {
        throw "WebView CSS pixels do not match the Office child window: $($measurement.Value)"
    }
    Write-Host "PASS browser pixel ratio $($metrics.dpr) matches the Office child window, without double DPI scaling."
    $before = (Get-Content -LiteralPath $log.FullName -Raw).Length
    $sheet.Range('A3').Formula2 = '=DuckQuery("SELECT i::DOUBLE x, cos(i/10.0) y FROM range(100) r(i)")'
    [void](Wait-XlDuckCell $sheet 'A4' -Until { param($v) $v -is [string] -and $v -cmatch '^xlduck://[tfp]/[A-Z]{4}/[1-9][0-9]*$' })
    $watch.Restart()
    do {
        $updates = (Get-Content -LiteralPath $log.FullName -Raw).Substring($before)
        if ($updates -match 'SetState: plot') { break }
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw 'The selected plot did not update after its dependency changed.' }
        Start-Sleep -Milliseconds 100
    } while ($true)
    if ($updates -match 'SetState: empty') { throw 'A transient dependency error cleared the previous chart.' }
    Write-Host 'PASS the selected chart updates without reselection and retains its previous state during dependency calculation.'
    function Wait-Preview([int]$after, [string]$pattern) {
        $wait = [Diagnostics.Stopwatch]::StartNew()
        do {
            $updates = (Get-Content -LiteralPath $log.FullName -Raw).Substring($after)
            if ($updates -match $pattern) { return }
            if ($wait.Elapsed.TotalSeconds -gt 15) { throw "Missing preview update: $pattern" }
            Start-Sleep -Milliseconds 100
        } while ($true)
    }
    function Select-Preview([string]$cell, [string]$kind, [string]$handle = '') {
        $after = (Get-Content -LiteralPath $log.FullName -Raw).Length
        $sheet.Range('Z100').Select() | Out-Null
        Wait-Preview $after 'SetState: empty'
        $after = (Get-Content -LiteralPath $log.FullName -Raw).Length
        $sheet.Range($cell).Select() | Out-Null
        Wait-Preview $after ('SetState: ' + $kind + ', Handle=' + [regex]::Escape($handle))
    }
    $sheet.Range('C3').Formula2 = '=DuckOut(A3)'
    [void](Wait-XlDuckCell $sheet 'C102' -Until { param($v) $v -eq 98 })
    $sourceHandle = [string]$sheet.Range('A3').Value2
    Select-Preview 'C3' 'table' $sourceHandle
    Select-Preview 'D10' 'table' $sourceHandle
    Write-Host 'PASS DuckOut spill header and interior select the original retained table.'
    $after = (Get-Content -LiteralPath $log.FullName -Raw).Length
    $sheet.Range('A3').Formula2 = '=DuckQuery("SELECT i::DOUBLE x, i::DOUBLE y FROM range(2) r(i)")'
    [void](Wait-XlDuckCell $sheet 'C5' -Until { param($v) $v -eq 1 })
    Wait-Preview $after 'SetState: empty'
    $sourceHandle = [string]$sheet.Range('A3').Value2
    Select-Preview 'D4' 'table' $sourceHandle
    $sheet.Range('C3').Cut($sheet.Range('F3')) | Out-Null
    [void](Wait-XlDuckCell $sheet 'F5' -Until { param($v) $v -eq 1 })
    Select-Preview 'G4' 'table' $sourceHandle
    $sheet.Range('F3').ClearContents() | Out-Null
    Select-Preview 'G4' 'empty'
    $sheet.Range('I3').Formula2 = '=SEQUENCE(3,2)'
    Select-Preview 'J4' 'empty'
    Write-Host 'PASS shrinking, moving and removing spills cannot leave a stale preview; ordinary Excel spills stay unrelated.'
    $sheet.Range('L3').Formula2 = '=DuckQueryOut("SELECT 7 x, 8 y")'
    [void](Wait-XlDuckCell $sheet 'M4' -Until { param($v) $v -eq 8 })
    Select-Preview 'M4' 'table'
    Write-Host 'PASS direct DuckQueryOut spills also expose their retained table.'
    if ($InspectionSeconds -gt 0) {
        Write-Host "Inspect owned Excel PID $($session.ProcessId); closing in $InspectionSeconds seconds."
        Start-Sleep -Seconds $InspectionSeconds
    }
    $excel.Run('DuckTogglePreview') | Out-Null
} finally {
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($session) { Close-XlDuckExcelSession $session }
}
