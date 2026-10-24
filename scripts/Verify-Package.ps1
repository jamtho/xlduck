# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$Directory)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($Directory)
$manifest = Get-Content -LiteralPath (Join-Path $root 'manifest.json') -Raw | ConvertFrom-Json
$seen = @{}
foreach ($entry in $manifest) {
    $path = [IO.Path]::GetFullPath((Join-Path $root $entry.path))
    if (-not $path.StartsWith($root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Manifest path escapes package directory.' }
    if ($seen.ContainsKey($entry.path)) { throw "Duplicate manifest path: $($entry.path)" }
    $seen[$entry.path] = $true
    $file = Get-Item -LiteralPath $path
    if ($file.Length -ne $entry.bytes -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $entry.sha256) { throw "Package file mismatch: $($entry.path)" }
}
foreach ($file in Get-ChildItem -LiteralPath $root -File -Recurse) {
    $relative = $file.FullName.Substring($root.Length+1).Replace('\','/')
    if ($relative -ne 'manifest.json' -and -not $seen.ContainsKey($relative)) { throw "Unmanifested package file: $relative" }
    if ($relative -match '(win-x86|win-arm64|\.pdb$)' -or $relative -eq 'XlDuck-AddIn.xll') { throw "Unintended package architecture/debug file: $relative" }
}
foreach ($required in @('XlDuck-AddIn64.xll','worker/XlDuck.Worker.exe','worker/runtimes/win-x64/native/duckdb.dll','runtimes/win-x64/native/WebView2Loader.dll','readers/httpfs.duckdb_extension','readers/aws.duckdb_extension','preview/preview.js','preview/vega.min.js','preview/vega-interpreter.min.js','licenses/nuget-packages.json','build-info.json','scripts/Get-Diagnostics.ps1')) {
    if (-not $seen.ContainsKey($required)) { throw "Missing required package file: $required" }
}
Write-Host "PASS package manifest: $($manifest.Count) files; x64 host, worker, readers and preview present."
