# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# File-only release regressions. Never starts, attaches to, or closes Excel.
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$verify = Join-Path $repo 'scripts/Verify-Package.ps1'
$diagnostics = Join-Path $repo 'scripts/Get-Diagnostics.ps1'
$source = (Resolve-Path -LiteralPath $PackageDirectory).Path
$sourceManifest = Get-Content -LiteralPath (Join-Path $source 'manifest.json') -Raw | ConvertFrom-Json
$testRoot = [IO.Path]::GetFullPath((Join-Path $repo ('artifacts/package-tests/' + [guid]::NewGuid().ToString('N'))))
[void][IO.Directory]::CreateDirectory($testRoot)
Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
$passed = 0; $failed = 0
$junctions = [Collections.Generic.List[string]]::new()
function Check([string]$Name, [scriptblock]$Action) {
    try { & $Action; $script:passed++; Write-Host "PASS $Name" }
    catch { $script:failed++; Write-Host "FAIL $Name : $($_.Exception.Message)" }
}
function Reject([scriptblock]$Action, [string]$Code) {
    $rejected = $false
    try { & $Action | Out-Null }
    catch {
        $rejected = $true
        if ($Code -and $_.Exception.Data['XlDuckPackageIssue'] -ne $Code) {
            throw "Expected package issue $Code; got $($_.Exception.Data['XlDuckPackageIssue'])."
        }
    }
    if (-not $rejected) { throw 'Invalid package was accepted.' }
}
function Save-Manifest([string]$Folder, $Entries) {
    ConvertTo-Json -InputObject @($Entries) -Depth 4 | Set-Content -LiteralPath (Join-Path $Folder 'manifest.json') -Encoding UTF8
}
function New-Fixture {
    $folder = Join-Path $testRoot ([guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($folder)
    $entries = foreach ($entry in $sourceManifest) {
        $file = [IO.Path]::GetFullPath((Join-Path $folder $entry.path))
        if (-not $file.StartsWith($folder + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture source path escaped its directory.' }
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($file))
        [IO.File]::WriteAllText($file, 'fixture ' + $entry.path)
        if ($entry.path -eq 'build-info.json') {
            [IO.File]::WriteAllText($file, '{"sourceCommit":"0123456789012345678901234567890123456789","trackedChanges":false,"untrackedFilesPresent":false}')
        }
        [pscustomobject]@{path=$entry.path;bytes=(Get-Item -LiteralPath $file).Length;sha256=(Get-FileHash -LiteralPath $file).Hash}
    }
    Save-Manifest $folder $entries
    return $folder
}
function Make-Archive([string]$Folder, [string]$Extra = '', [string]$Omit = '', [string]$Corrupt = '', [switch]$Duplicate) {
    $archive = [IO.Compression.ZipFile]::Open(($Folder + '.zip'), [IO.Compression.ZipArchiveMode]::Create)
    try {
        $prefix = (Split-Path $Folder -Leaf) + '/'
        foreach ($file in Get-ChildItem -LiteralPath $Folder -Recurse -File) {
            $relative = $file.FullName.Substring($Folder.Length+1).Replace('\','/')
            if ($relative -eq $Omit) { continue }
            if ($relative -eq $Corrupt) {
                $entry = $archive.CreateEntry($prefix + $relative)
                $writer = [IO.StreamWriter]::new($entry.Open())
                try { $writer.Write('corrupt') } finally { $writer.Dispose() }
            } else { [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $prefix + $relative) }
        }
        if ($Extra) { [void]$archive.CreateEntry($Extra.Replace('{root}', $prefix.TrimEnd('/'))) }
        if ($Duplicate) { [void]$archive.CreateEntry($prefix + 'XlDuck.dll') }
    } finally { $archive.Dispose() }
}
# Deterministic runtime discovery with deliberately private path text to redact.
$runtimeProbe = @{Fails = $false}
function dotnet {
    if ($runtimeProbe.Fails) { throw 'private-source-token: runtime failed at C:\private-source-token' }
    'Microsoft.WindowsDesktop.App 10.0.13 [C:\private-source-token\dotnet]'; $global:LASTEXITCODE = 0
}
try {
    Check 'complete folder and ZIP verify' {
        $folder = New-Fixture; Make-Archive $folder
        $null = & $verify -Directory $folder 6>&1
    }
    foreach ($required in @('XlDuck.dll','XlDuck.Core.dll','XlDuck-AddIn64.dna','XlDuck.runtimeconfig.json',
        'worker/XlDuck.Worker.dll','worker/XlDuck.Worker.deps.json','worker/XlDuck.Engine.dll','worker/Apache.Arrow.Scalars.dll',
        'preview/preview.html','preview/vega-lite.min.js','preview/leaflet.js','scripts/Verify-Package.ps1',
        'docs/INSTALL.md','docs/AGENT-INTEGRATION.md','docs/RELEASE-NOTES.md','docs/agent-exploration.js',
        'scripts/Setup.ps1','scripts/Setup.Core.ps1','START-HERE.html','Install.cmd',
        'agent/xlduck/SKILL.md','agent/xlduck-skill.zip','docs/AGENT-WORKFLOW.md',
        'THIRD-PARTY-NOTICES.md','third-party/notices.json','third-party/notices/duckdb-httpfs/LICENSE',
        'examples/research-demo/XlDuck Research Demo.xlsx','docs/assets/research-chart.jpg','docs/assets/research-table.jpg')) {
        Check "omitted required file is rejected even when its manifest entry is also missing: $required" {
            $folder = New-Fixture
            Remove-Item -LiteralPath (Join-Path $folder $required)
            $entries = Get-Content -LiteralPath (Join-Path $folder 'manifest.json') -Raw | ConvertFrom-Json
            Save-Manifest $folder @($entries | Where-Object path -ne $required)
            Reject { & $verify -Directory $folder 6>&1 } 'missing_required_file'
        }
    }
    Check 'changed file contents fail hash verification' {
        $folder = New-Fixture; [IO.File]::WriteAllText((Join-Path $folder 'XlDuck.dll'), 'damaged')
        Reject { & $verify -Directory $folder 6>&1 } 'file_mismatch'
    }
    Check 'missing listed file is reported separately' {
        $folder = New-Fixture; Remove-Item -LiteralPath (Join-Path $folder 'XlDuck.dll')
        Reject { & $verify -Directory $folder 6>&1 } 'missing_file'
    }
    Check 'unlisted files are rejected' {
        $folder = New-Fixture; [IO.File]::WriteAllText((Join-Path $folder 'extra.dll'), 'extra')
        Reject { & $verify -Directory $folder 6>&1 } 'unexpected_file'
    }
    foreach ($alias in @('preview/../XlDuck.dll','./XlDuck.dll','preview\preview.js','../outside.txt','C:/outside.txt')) {
        Check "noncanonical manifest path is rejected: $alias" {
            $folder = New-Fixture
            $entries = Get-Content -LiteralPath (Join-Path $folder 'manifest.json') -Raw | ConvertFrom-Json
            $item = $entries | Where-Object path -eq 'XlDuck.dll'
            Save-Manifest $folder ($entries + [pscustomobject]@{path=$alias;bytes=$item.bytes;sha256=$item.sha256})
            Reject { & $verify -Directory $folder 6>&1 } 'invalid_path'
        }
    }
    Check 'duplicate paths differing only by case are rejected' {
        $folder = New-Fixture
        $entries = Get-Content -LiteralPath (Join-Path $folder 'manifest.json') -Raw | ConvertFrom-Json
        $item = $entries | Where-Object path -eq 'XlDuck.dll'
        Save-Manifest $folder ($entries + [pscustomobject]@{path='xlduck.DLL';bytes=$item.bytes;sha256=$item.sha256})
        Reject { & $verify -Directory $folder 6>&1 } 'invalid_manifest'
    }
    Check 'malformed manifest metadata is rejected' {
        $folder = New-Fixture
        $entries = Get-Content -LiteralPath (Join-Path $folder 'manifest.json') -Raw | ConvertFrom-Json
        $entries[0].bytes = -1; Save-Manifest $folder $entries
        Reject { & $verify -Directory $folder 6>&1 } 'invalid_manifest'
    }
    foreach ($badJson in @('', '[]', '{"path":"XlDuck.dll"}', 'invalid JSON')) {
        Check "empty or invalid manifest JSON is rejected: '$badJson'" {
            $folder = New-Fixture
            [IO.File]::WriteAllText((Join-Path $folder 'manifest.json'), $badJson)
            Reject { & $verify -Directory $folder 6>&1 } 'invalid_manifest'
        }
    }
    Check 'junctions are rejected before package traversal' {
        $folder = New-Fixture; $destination = New-Fixture
        $link = Join-Path $folder 'linked'
        [void](New-Item -ItemType Junction -Path $link -Target $destination)
        $junctions.Add($link)
        Reject { & $verify -Directory $folder 6>&1 } 'linked_path'
        if (-not (Test-Path -LiteralPath (Join-Path $destination 'XlDuck.dll'))) { throw 'The link target was changed.' }
    }
    foreach ($entry in @('{root}/../../outside/','{root}/unexpected/')) {
        Check "unmanifested ZIP directory is rejected: $entry" {
            $folder = New-Fixture; Make-Archive $folder -Extra $entry
            Reject { & $verify -Directory $folder 6>&1 } 'invalid_archive'
        }
    }
    Check 'duplicate ZIP file is rejected' {
        $folder = New-Fixture; Make-Archive $folder -Duplicate
        Reject { & $verify -Directory $folder 6>&1 } 'invalid_archive'
    }
    Check 'missing ZIP file is rejected' {
        $folder = New-Fixture; Make-Archive $folder -Omit 'worker/XlDuck.Worker.dll'
        Reject { & $verify -Directory $folder 6>&1 } 'archive_mismatch'
    }
    Check 'changed ZIP contents are rejected' {
        $folder = New-Fixture; Make-Archive $folder -Corrupt 'preview/preview.js'
        Reject { & $verify -Directory $folder 6>&1 } 'archive_mismatch'
    }
    Check 'diagnostics identify the source build without collecting private paths or contents' {
        $folder = New-Fixture
        $json = & $diagnostics -PackageDirectory $folder
        $report = $json | ConvertFrom-Json
        if ($report.packageManifest -ne 'passed' -or $null -ne $report.packageIssue) { throw 'Valid package diagnostics failed.' }
        if ($report.sourceBuild.sourceCommit -ne '0123456789012345678901234567890123456789') { throw 'Source commit is absent.' }
        if ($report.availableDotnetRuntimes.Count -ne 1) { throw 'Runtime discovery failed.' }
        if ($json -match 'private-source-token' -or $json.Contains($folder) -or $json.Contains($folder.Replace('\','\\'))) { throw 'Diagnostics contain a private path.' }
    }
    Check 'damaged package diagnostics give a safe reason and remain valid JSON' {
        $folder = New-Fixture; [IO.File]::WriteAllText((Join-Path $folder 'XlDuck.dll'), 'private-source-token')
        $json = & $diagnostics -PackageDirectory $folder
        $report = $json | ConvertFrom-Json
        if ($report.packageManifest -ne 'failed' -or $report.packageIssue -ne 'file_mismatch') { throw 'Expected a safe package failure reason.' }
        if ($json -match 'private-source-token' -or $json.Contains($folder)) { throw 'Failure diagnostics exposed private data.' }
    }
    Check 'missing package diagnostics remain readable' {
        $report = & $diagnostics -PackageDirectory (Join-Path $testRoot 'missing') | ConvertFrom-Json
        if ($report.packageIssue -ne 'missing_directory' -or $null -ne $report.sourceBuild) { throw 'Missing package was not identified.' }
        if (@($report.binaries | Where-Object present).Count) { throw 'Missing binaries were reported present.' }
    }
    Check 'runtime discovery failure is reported without private error text' {
        $folder = New-Fixture
        $runtimeProbe.Fails = $true
        try { $json = & $diagnostics -PackageDirectory $folder }
        finally { $runtimeProbe.Fails = $false }
        $report = $json | ConvertFrom-Json
        if ($report.runtimeDiscovery -ne 'failed' -or $report.availableDotnetRuntimes.Count -ne 0) { throw 'Runtime failure was hidden.' }
        if ($json -match 'private-source-token') { throw 'Runtime error exposed private text.' }
    }
} finally {
    # Remove only this test's known junctions, never their target contents.
    foreach ($link in $junctions) {
        if (-not [IO.Path]::GetFullPath($link).StartsWith($testRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe link cleanup.' }
        if (Test-Path -LiteralPath $link) { [IO.Directory]::Delete($link) }
    }
    $allowed = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts/package-tests')).TrimEnd('\') + '\'
    if (-not $testRoot.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -or
        ((Get-Item -LiteralPath $testRoot).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Unsafe fixture cleanup.' }
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
Write-Host "$passed passed, $failed failed."
if ($failed) { exit 1 }
