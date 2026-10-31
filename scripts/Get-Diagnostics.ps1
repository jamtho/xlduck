# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$package = [IO.Path]::GetFullPath($PackageDirectory)
$verification = 'failed'
$packageIssue = $null
try {
    $null = & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $package 6>&1
    $verification = 'passed'
} catch {
    # Only emit our fixed diagnostic codes, never exception messages or paths.
    $knownIssues = @('missing_directory','missing_manifest','invalid_manifest','invalid_path','linked_path',
        'missing_file','file_mismatch','unexpected_file','unexpected_architecture','missing_required_file','invalid_archive','archive_mismatch')
    $issue = $_.Exception.Data['XlDuckPackageIssue']
    $packageIssue = if ($issue -cin $knownIssues) { $issue } else { 'unavailable' }
}
$sourceBuild = $null
try {
    $build = Get-Content -LiteralPath (Join-Path $package 'build-info.json') -Raw | ConvertFrom-Json
    if ($build.sourceCommit -is [string] -and $build.sourceCommit -cmatch '^(?:[0-9a-f]{40}|[0-9a-f]{64})$' -and
        $build.trackedChanges -is [bool] -and $build.untrackedFilesPresent -is [bool]) {
        $sourceBuild = [ordered]@{sourceCommit=$build.sourceCommit;trackedChanges=$build.trackedChanges;untrackedFilesPresent=$build.untrackedFilesPresent}
    }
} catch { } # Invalid/missing metadata is represented as null, not private error text.
$runtimes = @()
$runtimeDiscovery = 'not_found'
if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    try {
        $runtimeLines = @(dotnet --list-runtimes 2>$null)
        if ($LASTEXITCODE -ne 0) { throw 'Runtime discovery failed.' }
        $runtimes = @($runtimeLines | ForEach-Object {
            if ($_ -match '^(Microsoft\.[A-Za-z.]+) ([0-9A-Za-z.+-]+) \[') {
                [pscustomobject]@{name=$Matches[1];version=$Matches[2]}
            }
        })
        $runtimeDiscovery = 'available'
    } catch { $runtimeDiscovery = 'failed' }
}
$binaries = @('XlDuck.dll','XlDuck-AddIn64.xll','worker/XlDuck.Worker.exe') | ForEach-Object {
    $file = Join-Path $package $_
    $present = [IO.File]::Exists($file)
    $version = $null
    if ($present) {
        try { $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($file).FileVersion }
        catch { } # Still return a valid support report if a binary is unreadable.
    }
    [pscustomobject]@{
        name=$_
        present=$present
        version=$version
    }
}
[ordered]@{
    schemaVersion=2
    capturedUtc=[DateTime]::UtcNow.ToString('o')
    windowsVersion=[Environment]::OSVersion.Version.ToString()
    operatingSystem64Bit=[Environment]::Is64BitOperatingSystem
    diagnosticProcess64Bit=[Environment]::Is64BitProcess
    packageManifest=$verification
    packageIssue=$packageIssue
    sourceBuild=$sourceBuild
    binaries=@($binaries)
    availableDotnetRuntimes=$runtimes
    runtimeDiscovery=$runtimeDiscovery
    note='No workbook content, paths, settings, credentials or logs are collected. DuckStatus() reports the runtime actually loaded inside Excel.'
} | ConvertTo-Json -Depth 5
