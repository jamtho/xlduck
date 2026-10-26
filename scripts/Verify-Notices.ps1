# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([string]$Root = (Join-Path $PSScriptRoot '..'))
$ErrorActionPreference = 'Stop'
$folder = [IO.Path]::GetFullPath((Join-Path $Root 'third-party/notices'))
$entries = Get-Content -LiteralPath (Join-Path $Root 'third-party/notices.json') -Raw | ConvertFrom-Json
if ($entries.Count -eq 0) { throw 'Third-party notice inventory is empty.' }
$seen = @{}
foreach ($entry in $entries) {
    if ($entry.path -notmatch '^[A-Za-z0-9_. /-]+$' -or $entry.path -match '(^|/)\.\.?(/|$)' -or $seen.ContainsKey($entry.path)) {
        throw 'Invalid or duplicate third-party notice path.'
    }
    $path = [IO.Path]::GetFullPath((Join-Path $folder $entry.path))
    if (-not $path.StartsWith($folder + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Notice path escaped its directory.' }
    if ($entry.source -notmatch '^https://raw\.githubusercontent\.com/[^/]+/[^/]+/[0-9a-f]{40}/' -or
        $entry.sha256 -notmatch '^[0-9A-Fa-f]{64}$' -or
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ine $entry.sha256) { throw "Unpinned or changed notice: $($entry.path)" }
    $seen[$entry.path] = $true
}
if (@(Get-ChildItem -LiteralPath $folder -Recurse -File).Count -ne $entries.Count) { throw 'Unlisted third-party notice file.' }
Write-Host "PASS $($entries.Count) pinned third-party notices."
