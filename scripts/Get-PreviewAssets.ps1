# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Explicit build dependency acquisition; Excel never downloads executable renderer code.
param([string]$Destination = (Join-Path $PSScriptRoot '../artifacts/preview-assets'))
$ErrorActionPreference = 'Stop'
$target = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($target)
$packages = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'preview-assets.json') -Raw | ConvertFrom-Json
foreach ($package in $packages) {
    $temporary = Join-Path $target ('.download-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($temporary)
    try {
        $archive = Join-Path $temporary 'package.tgz'
        Invoke-WebRequest -UseBasicParsing -Uri "https://registry.npmjs.org/$($package.name)/-/$($package.name)-$($package.version).tgz" -OutFile $archive
        $sha = [Security.Cryptography.SHA512]::Create()
        $stream = [IO.File]::OpenRead($archive)
        try { $actual = 'sha512-' + [Convert]::ToBase64String($sha.ComputeHash($stream)) }
        finally { $stream.Dispose(); $sha.Dispose() }
        if ($actual -cne $package.integrity) { throw "Package integrity mismatch: $($package.name)" }
        # Extract only the named, integrity-verified package members, never an arbitrary archive tree.
        foreach ($file in $package.files.PSObject.Properties) {
            & tar -xzf $archive -C $temporary ('package/' + $file.Name)
            if ($LASTEXITCODE -ne 0) { throw "Missing asset: $($package.name)/$($file.Name)" }
            $destinationFile = Join-Path $target $file.Value
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destinationFile))
            Copy-Item -LiteralPath (Join-Path $temporary ('package/' + $file.Name)) -Destination $destinationFile -Force
        }
        Write-Host "Verified $($package.name) $($package.version)"
    } finally {
        $owned = [IO.Path]::GetFullPath($temporary)
        if (-not $owned.StartsWith($target.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe temporary cleanup path.' }
        if (Test-Path -LiteralPath $owned) { Remove-Item -LiteralPath $owned -Recurse -Force }
    }
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'preview-assets.json') -Destination (Join-Path $target 'versions.json') -Force
# These two npm distributions omit the repository-wide licence.
$license = Join-Path $target 'licenses/vega.txt'
Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/vega/vega/v5.33.1/LICENSE' -OutFile $license
if ((Get-FileHash -LiteralPath $license -Algorithm SHA256).Hash -ne '63727832AAF62004A2B249C933E327F3C90CAF49E41A72F4BF436EDF632CBDA8') { throw 'Vega licence hash mismatch.' }
Copy-Item -LiteralPath $license -Destination (Join-Path $target 'licenses/vega-interpreter.txt') -Force
