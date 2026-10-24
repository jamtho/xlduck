# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([switch]$NoBuild, [switch]$NoRestore, [switch]$WithExcel)
$ErrorActionPreference = 'Stop'
Push-Location (Join-Path $PSScriptRoot '..')
try {
    if (-not $NoBuild) {
        $projects = @('XlDuck.Worker', 'tests/XlDuck.Core.Specs', 'tests/XlDuck.Engine.Specs')
        if ($WithExcel) { $projects += 'tests/XlDuck.TestAddIn' }
        foreach ($project in $projects) {
            $options = @('build', $project, '--nologo', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false')
            if ($NoRestore) { $options += '--no-restore' }
            & dotnet @options
            if ($LASTEXITCODE -ne 0) { throw "Build failed: $project" }
        }
    }
    & dotnet run --project tests/XlDuck.Core.Specs --no-build --no-restore
    if ($LASTEXITCODE -ne 0) { throw 'Core specifications failed.' }
    & dotnet run --project tests/XlDuck.Engine.Specs --no-build --no-restore
    if ($LASTEXITCODE -ne 0) { throw 'Engine specifications failed.' }
    & dotnet run --project tests/XlDuck.Engine.Specs --no-build --no-restore -- --worker
    if ($LASTEXITCODE -ne 0) { throw 'Worker specifications failed.' }
    if ($WithExcel) {
        & powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-NewCore.ps1 -NoBuild
        if ($LASTEXITCODE -ne 0) { throw 'New-core Excel integration failed.' }
    }
} finally { Pop-Location }
