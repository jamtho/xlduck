# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# File/registry operations are separate from machine discovery so lifecycle tests
# can use an isolated directory and registry key, never the user's Excel options.
function Assert-XlDuckPlainPath([string]$Path) {
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            if ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Setup does not follow links or junctions: $current"
            }
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
}

function Get-XlDuckInstallState([string]$Root) {
    Assert-XlDuckPlainPath $Root
    $file = Join-Path $Root 'state.json'
    if (-not (Test-Path -LiteralPath $file)) {
        return [pscustomobject]@{schema=1;product='XlDuck';active=$null;previous=$null;versions=@()}
    }
    $state = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    if ($state.schema -ne 1 -or $state.product -cne 'XlDuck') { throw 'Unrecognized installation state; no changes made.' }
    $ids = @($state.versions)
    foreach ($id in $ids) {
        if ($id -isnot [string] -or $id -cnotmatch '^package-[0-9a-f]{16}$') { throw 'Invalid installed version identifier.' }
    }
    if (($state.active -and $state.active -notin $ids) -or ($state.previous -and $state.previous -notin $ids)) {
        throw 'Installed version history is inconsistent.'
    }
    return $state
}

function Save-XlDuckInstallState([string]$Root, $State) {
    $destination = Join-Path $Root 'state.json'
    $temporary = Join-Path $Root ('state-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($temporary, ($State | ConvertTo-Json -Depth 5))
        if ([IO.File]::Exists($destination)) { [IO.File]::Replace($temporary, $destination, ($destination + '.previous')) }
        else { [IO.File]::Move($temporary, $destination) }
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
}

function Get-XlDuckStartupEntries($Key) {
    if ($null -eq $Key) { return }
    foreach ($name in $Key.GetValueNames() | Where-Object { $_ -match '^OPEN(?:[1-9][0-9]*)?$' } |
        Sort-Object { if ($_ -ieq 'OPEN') { 0 } else { [int]$_.Substring(4) } }) {
        [pscustomobject]@{name=$name;value=$Key.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames);kind=$Key.GetValueKind($name)}
    }
}

function Get-XlDuckStartupPlan($Entries, [string[]]$ManagedPaths, [string]$Target) {
    $values = [Collections.Generic.List[object]]::new()
    $inserted = $false
    foreach ($entry in $Entries) {
        $path = $null
        if ($entry.value -is [string]) {
            $path = $entry.value.Trim() -replace '(?i)^/R\s+', ''
            $path = $path.Trim('"')
        }
        if ($path -and $path -in $ManagedPaths) {
            if ($Target -and -not $inserted) {
                $values.Add([pscustomobject]@{value=('/R "' + $Target + '"');kind=[Microsoft.Win32.RegistryValueKind]::String})
                $inserted = $true
            }
        } else {
            if ($Target -and $path -match '(?i)(?:^|[\\/])XlDuck[^\\/]*\.xll$') {
                throw 'Another XlDuck build is registered outside this installer. Uncheck it in Excel Add-ins, fully close Excel, then run setup again.'
            }
            $values.Add($entry)
        }
    }
    if ($Target -and -not $inserted) {
        $values.Add([pscustomobject]@{value=('/R "' + $Target + '"');kind=[Microsoft.Win32.RegistryValueKind]::String})
    }
    for ($i=0; $i -lt $values.Count; $i++) {
        [pscustomobject]@{name=$(if ($i -eq 0) {'OPEN'} else {'OPEN' + $i});value=$values[$i].value;kind=$values[$i].kind}
    }
}

function Set-XlDuckStartupEntries($Key, $Entries) {
    # Excel requires a contiguous OPEN, OPEN1, ... list. Keep every unrelated
    # entry's value, type and relative load order; never touch other options.
    foreach ($name in @($Key.GetValueNames() | Where-Object { $_ -match '^OPEN(?:[1-9][0-9]*)?$' })) { $Key.DeleteValue($name) }
    foreach ($entry in $Entries) { $Key.SetValue($entry.name, $entry.value, $entry.kind) }
    $Key.Flush()
}

function Assert-XlDuckExcelClosed {
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) {
        throw 'Save your work and close all Excel windows, then run setup again. Setup will not close Excel for you.'
    }
}

function Invoke-XlDuckInstall {
    param([ValidateSet('Install','Rollback','Uninstall')][string]$Action,
        [string]$Source, [string]$Root, [string]$OptionsKey,
        [scriptblock]$AssertClosed = { Assert-XlDuckExcelClosed })
    & $AssertClosed
    $Root = [IO.Path]::GetFullPath($Root)
    $state = Get-XlDuckInstallState $Root
    $targetId = $null; $target = $null
    if ($Action -eq 'Rollback') {
        if (-not $state.previous) { throw 'There is no previous installed version to restore.' }
        $Source = Join-Path $Root ('versions/' + $state.previous)
    }
    if ($Action -ne 'Uninstall') {
        $Source = [IO.Path]::GetFullPath($Source)
        & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $Source
        # No implicit trust changes. Downloaded packages must be unblocked by
        # the user after reviewing their origin, before setup copies them.
        foreach ($file in Get-ChildItem -LiteralPath $Source -Recurse -File) {
            $zone = Get-Content -LiteralPath $file.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue
            if ($zone -match '^ZoneId=[34]$') { throw 'This package is marked as downloaded. Review its origin, unblock the ZIP in Properties and extract again. See docs/INSTALL.md.' }
        }
        $targetId = 'package-' + (Get-FileHash -LiteralPath (Join-Path $Source 'manifest.json')).Hash.Substring(0,16).ToLowerInvariant()
        $target = Join-Path $Root ('versions/' + $targetId)
        Assert-XlDuckPlainPath $target
    }
    $managed = @($state.versions | ForEach-Object { Join-Path $Root ('versions/' + $_ + '/XlDuck-AddIn64.xll') })
    $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($OptionsKey, $true)
    try {
        $original = @(Get-XlDuckStartupEntries $key)
        $targetXll = if ($target) { Join-Path $target 'XlDuck-AddIn64.xll' } else { $null }
        $next = @(Get-XlDuckStartupPlan $original $managed $targetXll)
        if ($target) {
            if (Test-Path -LiteralPath $target) {
                & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $target
                if ((Get-FileHash -LiteralPath (Join-Path $target 'manifest.json')).Hash -ne
                    (Get-FileHash -LiteralPath (Join-Path $Source 'manifest.json')).Hash) { throw 'Version directory collision.' }
            } else {
                [void][IO.Directory]::CreateDirectory($target)
                foreach ($item in Get-ChildItem -LiteralPath $Source -Force) { Copy-Item -LiteralPath $item.FullName -Destination $target -Recurse }
                & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $target
            }
        }
        [void][IO.Directory]::CreateDirectory($Root)
        # Keep recovery information before changing the registration. Never
        # replay this snapshot during a later uninstall: Excel may have changed.
        $backup = Join-Path $Root ('startup-before-' + [guid]::NewGuid().ToString('N') + '.json')
        ConvertTo-Json -InputObject $original -Depth 5 | Set-Content -LiteralPath $backup -Encoding UTF8
        & $AssertClosed
        if (-not $key) { $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($OptionsKey) }
        # Refuse a stale write if another process changed the list during copy.
        if ((ConvertTo-Json -InputObject @(Get-XlDuckStartupEntries $key) -Compress) -ne
            (ConvertTo-Json -InputObject $original -Compress)) { throw 'Excel startup settings changed during setup. Run setup again.' }
        try {
            Set-XlDuckStartupEntries $key $next
            if ($targetId) {
                if ($state.active -ne $targetId) { $state.previous = $state.active }
                $state.active = $targetId
                $state.versions = @(@($state.versions) + $targetId | Select-Object -Unique)
            } else { $state.active = $null; $state.previous = $null }
            Save-XlDuckInstallState $Root $state
        } catch {
            Set-XlDuckStartupEntries $key $original
            throw
        }
    } finally { if ($key) { $key.Dispose() } }
    # Leave version directories in place on uninstall for recovery/manual removal.
    # This also avoids deleting a workbook or user-added file from the install tree.
    [pscustomobject]@{action=$Action;directory=$target;state=$state;root=$Root}
}

function Install-XlDuckLocalSkill([string]$Source, [string]$Destination) {
    Assert-XlDuckPlainPath $Destination
    $sourceFile = Join-Path $Source 'SKILL.md'
    if (-not [IO.File]::Exists($sourceFile)) { throw 'Packaged bootstrap skill is missing.' }
    $targetFile = Join-Path $Destination 'SKILL.md'
    Assert-XlDuckPlainPath $targetFile
    if (Test-Path -LiteralPath $Destination) {
        $items = @(Get-ChildItem -LiteralPath $Destination -Force)
        if ($items.Count -eq 1 -and $items[0].Name -ceq 'SKILL.md' -and
            (Get-FileHash -LiteralPath $sourceFile).Hash -eq (Get-FileHash -LiteralPath $targetFile).Hash) { return }
        throw 'An existing xlduck skill differs from this package. Keep or move your existing skill before installing; setup will not overwrite it.'
    }
    [void][IO.Directory]::CreateDirectory($Destination)
    Copy-Item -LiteralPath $sourceFile -Destination $targetFile
}
