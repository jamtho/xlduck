# Copyright (c) 2026 James Thompson
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

# Reload XlDuck add-in: cleanly close Excel, rebuild, relaunch with blank workbook

param(
    [switch]$NoBuild
)

$repoRoot = Split-Path -Parent $PSScriptRoot
$addInPath = Join-Path $repoRoot "XlDuck\bin\Debug\net8.0-windows\XlDuck-AddIn64.xll"
$projectPath = Join-Path $repoRoot "XlDuck\XlDuck.csproj"

# Step 1: Cleanly close Excel via COM
Write-Host "Closing Excel..." -ForegroundColor Cyan
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $excel.DisplayAlerts = $false  # Suppress save prompts

    # Close all workbooks without saving (iterate backwards to avoid collection modification issues)
    while ($excel.Workbooks.Count -gt 0) {
        $excel.Workbooks.Item(1).Close($false)
    }

    # Quit Excel
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Wait for Excel to fully exit
    $timeout = 10
    while ((Get-Process excel -ErrorAction SilentlyContinue) -and $timeout -gt 0) {
        Start-Sleep -Milliseconds 500
        $timeout--
    }

    if ($timeout -eq 0) {
        Write-Host "  Excel didn't close gracefully, forcing..." -ForegroundColor Yellow
        Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Milliseconds 500
    }

    Write-Host "  Excel closed" -ForegroundColor Green
} catch {
    Write-Host "  Excel not running or error: $_" -ForegroundColor Yellow
    # Force kill if COM failed
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

# Step 2: Build (unless -NoBuild)
if (-not $NoBuild) {
    Write-Host "Building..." -ForegroundColor Cyan
    $buildResult = & "C:\Program Files\dotnet\dotnet.exe" build $projectPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Build failed!" -ForegroundColor Red
        $buildResult | Write-Host
        exit 1
    }
    Write-Host "  Build succeeded" -ForegroundColor Green
}

# Step 3: Launch Excel via COM (registers with ROT properly)
Write-Host "Launching Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# Step 4: Create blank workbook first
Write-Host "Creating blank workbook..." -ForegroundColor Cyan
$excel.Workbooks.Add() | Out-Null
Write-Host "  Created blank workbook" -ForegroundColor Green

# Step 5: Register the XLL add-in
Write-Host "Loading add-in..." -ForegroundColor Cyan
try {
    # RegisterXLL loads the add-in for this session
    $registered = $excel.RegisterXLL($addInPath)
    if ($registered) {
        Write-Host "  Add-in loaded" -ForegroundColor Green
    } else {
        Write-Host "  Add-in registration returned false" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Error loading add-in: $_" -ForegroundColor Red
}

[Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Ready!" -ForegroundColor Green
