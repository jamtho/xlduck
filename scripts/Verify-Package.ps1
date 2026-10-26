# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$Directory)
$ErrorActionPreference = 'Stop'
function Fail-Package([string]$Code, [string]$Message) {
    $failure = [IO.InvalidDataException]::new($Message)
    $failure.Data['XlDuckPackageIssue'] = $Code
    throw $failure
}
function Assert-RelativePath([string]$Path, [string]$Code = 'invalid_path') {
    # One portable spelling per Windows file: no traversal, ADS, wildcard paths,
    # device names or aliases that normalize differently during ZIP extraction.
    if ([string]::IsNullOrEmpty($Path) -or $Path -match '[\\:\x00-\x1f<>"|?*]') {
        Fail-Package $Code 'Package paths must be canonical relative paths with forward slashes.'
    }
    foreach ($part in $Path.Split('/')) {
        if ($part -in @('', '.', '..') -or $part.EndsWith('.') -or $part.EndsWith(' ') -or
            $part -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') {
            Fail-Package $Code 'Package paths must not contain empty, reserved or traversal components.'
        }
    }
}
function Assert-OrdinaryItem($Item) {
    if ($Item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        Fail-Package 'linked_path' 'Packages must not contain symbolic links or junctions.'
    }
}
$root = [IO.Path]::GetFullPath($Directory)
if ($root -ne [IO.Path]::GetPathRoot($root)) { $root = $root.TrimEnd('\','/') }
$rootPrefixLength = $root.TrimEnd('\','/').Length + 1
if (-not (Test-Path -LiteralPath $root -PathType Container)) { Fail-Package 'missing_directory' 'Package directory is unavailable.' }
Assert-OrdinaryItem (Get-Item -LiteralPath $root -Force)
# Inspect each directory before descending, rather than following a junction while
# enumerating or hashing. Build inventory independently from the manifest.
$inventory = @{}
$pending = [Collections.Generic.Queue[string]]::new()
$pending.Enqueue($root)
while ($pending.Count) {
    foreach ($item in Get-ChildItem -LiteralPath $pending.Dequeue() -Force) {
        Assert-OrdinaryItem $item
        $relative = $item.FullName.Substring($rootPrefixLength).Replace('\','/')
        Assert-RelativePath $relative
        if ($item.PSIsContainer) { $pending.Enqueue($item.FullName) }
        else { $inventory[$relative] = $item }
    }
}
if (-not $inventory.ContainsKey('manifest.json')) { Fail-Package 'missing_manifest' 'Package manifest is missing.' }
try {
    $json = Get-Content -LiteralPath (Join-Path $root 'manifest.json') -Raw
    $manifest = $json | ConvertFrom-Json
} catch { Fail-Package 'invalid_manifest' 'Package manifest is not readable JSON.' }
if ($json -isnot [string] -or -not $json.TrimStart().StartsWith('[') -or @($manifest).Count -eq 0) {
    Fail-Package 'invalid_manifest' 'Package manifest must be a nonempty array.'
}
$seen = @{}
foreach ($entry in $manifest) {
    if ($entry.path -isnot [string] -or ($entry.bytes -isnot [int] -and $entry.bytes -isnot [long]) -or
        $entry.bytes -lt 0 -or $entry.sha256 -isnot [string] -or $entry.sha256 -notmatch '^[0-9A-Fa-f]{64}$') {
        Fail-Package 'invalid_manifest' 'Manifest entries require a path, nonnegative integer byte count and SHA-256 hash.'
    }
    Assert-RelativePath $entry.path
    if ($entry.path -ieq 'manifest.json' -or $seen.ContainsKey($entry.path)) {
        Fail-Package 'invalid_manifest' 'Manifest contains a duplicate or self-referential path.'
    }
    $seen[$entry.path] = $true
    if (-not $inventory.ContainsKey($entry.path)) { Fail-Package 'missing_file' "Package file is missing: $($entry.path)" }
    $file = $inventory[$entry.path]
    if ($file.Length -ne $entry.bytes -or (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -ine $entry.sha256) {
        Fail-Package 'file_mismatch' "Package file mismatch: $($entry.path)"
    }
}
foreach ($relative in $inventory.Keys) {
    if ($relative -ne 'manifest.json' -and -not $seen.ContainsKey($relative)) {
        Fail-Package 'unexpected_file' "Unmanifested package file: $relative"
    }
    if ($relative -match '(win-x86|win-arm64|\.pdb$)' -or $relative -eq 'XlDuck-AddIn.xll') {
        Fail-Package 'unexpected_architecture' "Unintended package architecture/debug file: $relative"
    }
}
# A manifest can be internally consistent yet omit a runtime dependency. Require
# the complete host, worker and UI baseline, independently of build enumeration.
$requiredFiles = @(
    'XlDuck-AddIn64.xll','XlDuck-AddIn64.dna','XlDuck-AddIn64.deps.json',
    'XlDuck.dll','XlDuck.deps.json','XlDuck.runtimeconfig.json','XlDuck.comhost.dll',
    'XlDuck.Core.dll','XlDuck.Client.dll','XlDuck.Host.dll',
    'Microsoft.Web.WebView2.Core.dll','Microsoft.Web.WebView2.WinForms.dll',
    'runtimes/win-x64/native/WebView2Loader.dll',
    'worker/XlDuck.Worker.exe','worker/XlDuck.Worker.dll','worker/XlDuck.Worker.deps.json',
    'worker/XlDuck.Worker.runtimeconfig.json','worker/XlDuck.Core.dll','worker/XlDuck.Engine.dll',
    'worker/DuckDB.NET.Bindings.dll','worker/DuckDB.NET.Data.dll',
    'worker/Apache.Arrow.dll','worker/Apache.Arrow.Scalars.dll','worker/runtimes/win-x64/native/duckdb.dll',
    'readers/httpfs.duckdb_extension','readers/aws.duckdb_extension',
    'preview/preview.html','preview/preview.css','preview/preview.js','preview/versions.json',
    'preview/vega.min.js','preview/vega-lite.min.js','preview/vega-embed.min.js','preview/vega-interpreter.min.js',
    'preview/leaflet.js','preview/leaflet.css','preview/sql-formatter.min.js',
    'preview/licenses/vega.txt','preview/licenses/vega-lite.txt','preview/licenses/vega-embed.txt',
    'preview/licenses/vega-interpreter.txt','preview/licenses/leaflet.txt','preview/licenses/sql-formatter.txt',
    'LICENSE','README.md','docs/FORMULAS.md','docs/RELEASE.md','docs/INSTALL.md','docs/AGENT-INTEGRATION.md','docs/agent-exploration.js','licenses/nuget-packages.json',
    'build-info.json','scripts/Get-Diagnostics.ps1','scripts/Verify-Package.ps1',
    'docs/AGENT-WORKFLOW.md','scripts/Setup.ps1','scripts/Setup.Core.ps1',
    'Install.cmd','Uninstall.cmd','Rollback.cmd','Install Codex Skill.cmd','START-HERE.html',
    'agent/xlduck/SKILL.md','agent/xlduck-skill.zip',
    'THIRD-PARTY-NOTICES.md','third-party/notices.json',
    'third-party/notices/duckdb-1.5.6/LICENSE','third-party/notices/duckdb-httpfs/LICENSE','third-party/notices/duckdb-aws/LICENSE',
    'docs/assets/xlduck-logo.png','docs/assets/research-chart.jpg','docs/assets/research-table.jpg',
    'examples/research-demo/XlDuck Research Demo.xlsx','examples/research-demo/README.md'
)
foreach ($required in $requiredFiles) {
    if (-not $seen.ContainsKey($required)) { Fail-Package 'missing_required_file' "Missing required package file: $required" }
}
if (Test-Path -LiteralPath ($root + '.zip')) {
    Assert-OrdinaryItem (Get-Item -LiteralPath ($root + '.zip') -Force)
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    try { $archive = [IO.Compression.ZipFile]::OpenRead($root + '.zip') }
    catch { Fail-Package 'invalid_archive' 'Package ZIP cannot be opened.' }
    try {
        $prefix = (Split-Path $root -Leaf) + '/'
        $expected = @{}
        foreach ($item in $manifest) { $expected[$prefix + $item.path] = $item }
        $manifestFile = $inventory['manifest.json']
        $expected[$prefix + 'manifest.json'] = @{bytes=$manifestFile.Length;sha256=(Get-FileHash -LiteralPath $manifestFile.FullName -Algorithm SHA256).Hash}
        $found = @{}
        foreach ($entry in $archive.Entries) {
            Assert-RelativePath $entry.FullName 'invalid_archive'
            # The builder emits files only. Reject directory entries too, so a
            # traversal path cannot evade validation simply by ending in '/'.
            if (-not $expected.ContainsKey($entry.FullName) -or $found.ContainsKey($entry.FullName)) {
                Fail-Package 'invalid_archive' 'ZIP has an unexpected or duplicate entry.'
            }
            $item = $expected[$entry.FullName]
            if ($entry.Length -ne $item.bytes) { Fail-Package 'archive_mismatch' 'ZIP entry size does not match its manifest.' }
            $stream = $entry.Open(); $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $hash = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')
                if ($hash -ine $item.sha256) { Fail-Package 'archive_mismatch' 'ZIP entry hash does not match its manifest.' }
            } finally { $stream.Dispose(); $sha.Dispose() }
            $found[$entry.FullName] = $true
        }
        if ($found.Count -ne $expected.Count) { Fail-Package 'archive_mismatch' 'ZIP omits manifested files.' }
    } finally { $archive.Dispose() }
    Write-Host "PASS distribution ZIP: $($found.Count) files including the manifest."
}
Write-Host "PASS package manifest: $(@($manifest).Count) files; complete x64 host, worker, readers and preview present."
