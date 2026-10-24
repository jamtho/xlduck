# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Development dependency acquisition only. Worksheet SQL never installs or downloads extensions.
param([string]$Destination = (Join-Path $PSScriptRoot '../artifacts/reader-extensions/v1.5.6/windows_amd64'))
$ErrorActionPreference = 'Stop'
$target = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($target)
foreach ($name in @('httpfs', 'aws')) {
    $extension = Join-Path $target "$name.duckdb_extension"
    $expectedHash = if ($name -eq 'httpfs') { '9B0B0202CFA34F94F12BD9B2698002735BBCD0C1BD5FDA2E85450583EFD18E94' } else { '008CB2809ADD3AFB7E5BCDF27655BE129099D1B35B52D3248969A5A055F4C627' }
    if (-not (Test-Path -LiteralPath $extension)) {
        $temporary = Join-Path $target ('.download-' + [Guid]::NewGuid().ToString('N'))
        $compressed = $temporary + '.gz'
        try {
            Invoke-WebRequest -UseBasicParsing -Uri "https://extensions.duckdb.org/v1.5.6/windows_amd64/$name.duckdb_extension.gz" -OutFile $compressed
            $inputStream = [IO.File]::OpenRead($compressed)
            try {
                $gzip = New-Object IO.Compression.GZipStream($inputStream, [IO.Compression.CompressionMode]::Decompress)
                try {
                    $outputStream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
                    try { $gzip.CopyTo($outputStream) }
                    finally { $outputStream.Dispose() }
                } finally { $gzip.Dispose() }
            } finally { $inputStream.Dispose() }
            if ((Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -ne $expectedHash) { throw "Downloaded reader extension hash mismatch: $name" }
            [IO.File]::Move($temporary, $extension)
        } finally {
            foreach ($ownedFile in @($temporary, $compressed)) {
                if (Test-Path -LiteralPath $ownedFile) { Remove-Item -LiteralPath $ownedFile }
            }
        }
    }
    $actualHash = Get-FileHash -LiteralPath $extension -Algorithm SHA256
    if ($actualHash.Hash -ne $expectedHash) { throw "Reader extension hash mismatch: $extension" }
    $actualHash | Select-Object Hash,Path
}
