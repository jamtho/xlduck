# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([string]$Destination, [string]$Dotnet = 'dotnet')
$ErrorActionPreference = 'Stop'
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $Destination) { $Destination = Join-Path $repo ('artifacts/packages/xlduck-0.2.0-dev-' + [guid]::NewGuid().ToString('N').Substring(0,8)) }
$package = [IO.Path]::GetFullPath($Destination)
if (Test-Path -LiteralPath $package) { throw 'Package destination already exists; choose a new directory.' }
if (Test-Path -LiteralPath ($package + '.zip')) { throw 'Package archive already exists; choose a new destination.' }
$build = Join-Path $repo ('artifacts/package-builds/' + [guid]::NewGuid().ToString('N'))
foreach ($asset in @('artifacts/preview-assets/vega.min.js','artifacts/reader-extensions/v1.5.6/windows_amd64/httpfs.duckdb_extension','artifacts/reader-extensions/v1.5.6/windows_amd64/aws.duckdb_extension')) {
    if (-not (Test-Path -LiteralPath (Join-Path $repo $asset))) { throw 'Acquire pinned readers and preview assets before packaging.' }
}
& $Dotnet build (Join-Path $repo 'XlDuck') -c Release --artifacts-path $build --nologo -m:1 /nodeReuse:false /p:UseSharedCompilation=false /p:RestoreLockedMode=true
if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }
$output = Join-Path $build 'bin/XlDuck/release'
[void][IO.Directory]::CreateDirectory($package)
# Explicit roots prevent stale development binaries and unintended architectures from shipping.
$rootFiles = @('XlDuck-AddIn64.xll','XlDuck-AddIn64.dna','XlDuck-AddIn64.deps.json',
    'XlDuck.dll','XlDuck.deps.json','XlDuck.runtimeconfig.json','XlDuck.comhost.dll',
    'XlDuck.Core.dll','XlDuck.Client.dll','XlDuck.Host.dll',
    'Microsoft.Web.WebView2.Core.dll','Microsoft.Web.WebView2.WinForms.dll')
foreach ($name in $rootFiles) { Copy-Item -LiteralPath (Join-Path $output $name) -Destination (Join-Path $package $name) }
foreach ($folder in @('worker','readers','preview','runtimes/win-x64/native')) {
    $destinationFolder = Join-Path $package $folder
    [void][IO.Directory]::CreateDirectory($destinationFolder)
    Get-ChildItem -LiteralPath (Join-Path $output $folder) | Where-Object { $_.Extension -ne '.pdb' } | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $destinationFolder -Recurse
    }
}
Copy-Item -LiteralPath (Join-Path $repo 'LICENSE'),(Join-Path $repo 'README.md') -Destination $package
[void][IO.Directory]::CreateDirectory((Join-Path $package 'docs'))
Copy-Item -LiteralPath (Join-Path $repo 'docs/FORMULAS.md'),(Join-Path $repo 'docs/RELEASE.md') -Destination (Join-Path $package 'docs')
[void][IO.Directory]::CreateDirectory((Join-Path $package 'scripts'))
Copy-Item -LiteralPath (Join-Path $repo 'scripts/Get-Diagnostics.ps1'),(Join-Path $repo 'scripts/Verify-Package.ps1') -Destination (Join-Path $package 'scripts')
$sourceCommit = & git -C $repo rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Could not identify source commit.' }
$trackedChanges = @(& git -C $repo status --porcelain --untracked-files=no).Count -gt 0
$untrackedFiles = @(& git -C $repo ls-files --others --exclude-standard).Count -gt 0
[ordered]@{sourceCommit=$sourceCommit;trackedChanges=$trackedChanges;untrackedFilesPresent=$untrackedFiles;builtUtc=[DateTime]::UtcNow.ToString('o');sdk=(& $Dotnet --version)} |
    ConvertTo-Json | Set-Content (Join-Path $package 'build-info.json')
$notices = Join-Path $package 'licenses'
[void][IO.Directory]::CreateDirectory($notices)
$packages = @{}
foreach ($assetsFile in @((Join-Path $build 'obj/XlDuck/project.assets.json'),(Join-Path $build 'obj/XlDuck.Worker/project.assets.json'))) {
    $assets = Get-Content -LiteralPath $assetsFile -Raw | ConvertFrom-Json
    foreach ($library in $assets.libraries.PSObject.Properties) {
        if ($library.Value.type -ne 'package') { continue }
        $packageDirectory = $null
        foreach ($root in $assets.packageFolders.PSObject.Properties.Name) {
            $candidate = Join-Path $root $library.Value.path
            if (Test-Path -LiteralPath $candidate) { $packageDirectory = $candidate; break }
        }
        if (-not $packageDirectory) { throw "Missing restored dependency $($library.Name)" }
        $noticeDirectory = Join-Path $notices ($library.Name.Replace('/','-'))
        [void][IO.Directory]::CreateDirectory($noticeDirectory)
        Get-ChildItem -LiteralPath $packageDirectory -File | Where-Object { $_.Name -match '(?i)license|notice|\.nuspec$' } | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination $noticeDirectory
        }
        $packages[$library.Name] = $library.Value.sha512
    }
}
$packages | ConvertTo-Json | Set-Content (Join-Path $notices 'nuget-packages.json')
$manifest = Get-ChildItem -LiteralPath $package -Recurse -File | ForEach-Object {
    [pscustomobject]@{path=$_.FullName.Substring($package.Length+1).Replace('\','/');bytes=$_.Length;sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash}
}
$manifest | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $package 'manifest.json')
# Windows PowerShell's Compress-Archive writes backslash entry names. Emit standard
# ZIP paths explicitly so extraction works consistently across archive tools.
Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::Open(($package + '.zip'), [IO.Compression.ZipArchiveMode]::Create)
try {
    $prefix = (Split-Path $package -Leaf) + '/'
    foreach ($file in Get-ChildItem -LiteralPath $package -Recurse -File) {
        $entry = $prefix + $file.FullName.Substring($package.Length+1).Replace('\','/')
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $entry, [IO.Compression.CompressionLevel]::Optimal)
    }
} finally { $archive.Dispose() }
Write-Host "Unsigned development package: $package"
Write-Host "Archive: $package.zip"
