# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Real file/registry lifecycle tests in disposable locations. No Excel automation,
# no edits to Office settings, no global skill installation or prerequisite install.
param([Parameter(Mandatory)][string]$PackageDirectory)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
. (Join-Path $repo 'scripts/Setup.Core.ps1')
$testRoot = Join-Path $repo ('artifacts/setup-tests/' + [guid]::NewGuid().ToString('N'))
$testKey = 'Software\XlDuck\SetupTests\' + [guid]::NewGuid().ToString('N')
[void][IO.Directory]::CreateDirectory($testRoot)
$key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($testKey)
$passed=0; $failed=0
function Check([string]$Name, [scriptblock]$Action) {
    try { & $Action; $script:passed++; Write-Host "PASS $Name" }
    catch { $script:failed++; Write-Host "FAIL $Name : $($_.Exception.Message)" }
}
function Assert([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Reject([scriptblock]$Action, [string]$Message) {
    try { & $Action | Out-Null } catch { if ($Message) { Assert ($_.Exception.Message -like ('*' + $Message + '*')) $_.Exception.Message }; return }
    throw 'Expected operation to be rejected.'
}
function Snapshot { ConvertTo-Json -InputObject @(Get-XlDuckStartupEntries $key) -Compress }
function New-Package([string]$Marker) {
    $folder = Join-Path $testRoot ('source-' + $Marker)
    [void][IO.Directory]::CreateDirectory($folder)
    $sourceEntries = Get-Content -LiteralPath (Join-Path $PackageDirectory 'manifest.json') -Raw | ConvertFrom-Json
    $manifest = foreach ($entry in $sourceEntries) {
        $file = Join-Path $folder $entry.path
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($file))
        [IO.File]::WriteAllText($file, ($Marker + ':' + $entry.path))
        [pscustomobject]@{path=$entry.path;bytes=(Get-Item -LiteralPath $file).Length;sha256=(Get-FileHash -LiteralPath $file).Hash}
    }
    ConvertTo-Json -InputObject @($manifest) | Set-Content -LiteralPath (Join-Path $folder 'manifest.json') -Encoding UTF8
    return $folder
}
$root = Join-Path $testRoot 'install'
$options = @{Root=$root;OptionsKey=$testKey;AssertClosed={}}
try {
    $source = New-Package 'first'; $second = New-Package 'second'
    $key.SetValue('OPEN','/R "C:\Other\first.xll"')
    $key.SetValue('OPEN3','%APPDATA%\Other\second.xla',[Microsoft.Win32.RegistryValueKind]::ExpandString)
    $key.SetValue('UnrelatedSetting',42,[Microsoft.Win32.RegistryValueKind]::DWord)
    Check 'install preserves unrelated entries, types and order and fills OPEN gaps' {
        $script:installed = Invoke-XlDuckInstall -Action Install -Source $source @options
        $entries = @(Get-XlDuckStartupEntries $key)
        Assert ($entries.Count -eq 3 -and $entries[0].name -eq 'OPEN' -and $entries[1].name -eq 'OPEN1' -and $entries[2].name -eq 'OPEN2') 'Startup list has gaps.'
        Assert ($entries[0].value -eq '/R "C:\Other\first.xll"' -and $entries[1].value -eq '%APPDATA%\Other\second.xla') 'Unrelated load order or values changed.'
        Assert ($entries[1].kind -eq 'ExpandString' -and $key.GetValue('UnrelatedSetting') -eq 42) 'Unrelated registry types/settings changed.'
        Assert (Test-Path -LiteralPath (Join-Path $installed.directory 'worker/XlDuck.Worker.exe')) 'Incomplete copied distribution.'
    }
    Check 'repeat install is idempotent and leaves one registration' {
        $before = Snapshot
        $again = Invoke-XlDuckInstall -Action Install -Source $source @options
        Assert ((Snapshot) -ceq $before) 'Repeat install changed registration.'
        Assert ($again.state.versions.Count -eq 1 -and -not $again.state.previous) 'Repeat install changed history.'
    }
    Check 'upgrade retains old distribution and records rollback' {
        $script:upgraded = Invoke-XlDuckInstall -Action Install -Source $second @options
        Assert ($upgraded.state.previous -eq $installed.state.active) 'Previous version was lost.'
        Assert (Test-Path -LiteralPath (Join-Path $installed.directory 'XlDuck.dll')) 'Upgrade removed old binaries.'
        Assert ((@(Get-XlDuckStartupEntries $key)[2].value).Contains($upgraded.directory)) 'Upgrade did not switch startup path.'
    }
    Check 'rollback switches back without copying credentials or deleting newer version' {
        $rolled = Invoke-XlDuckInstall -Action Rollback @options
        Assert ($rolled.state.active -eq $installed.state.active -and $rolled.state.previous -eq $upgraded.state.active) 'Rollback history incorrect.'
        Assert (Test-Path -LiteralPath (Join-Path $upgraded.directory 'XlDuck.dll')) 'Rollback removed newer binaries.'
    }
    Check 'Excel opening during copy prevents registration changes' {
        $before = Snapshot; $stateBefore = Get-Content -LiteralPath (Join-Path $root 'state.json') -Raw
        $gate = @{calls=0}
        Reject { Invoke-XlDuckInstall -Action Install -Source $second -Root $root -OptionsKey $testKey -AssertClosed { $gate.calls++; if ($gate.calls -eq 2) { throw 'Excel opened during setup' } } } 'Excel opened'
        Assert ((Snapshot) -ceq $before) 'Registration changed despite Excel opening.'
        Assert ((Get-Content -LiteralPath (Join-Path $root 'state.json') -Raw) -ceq $stateBefore) 'State changed despite Excel opening.'
    }
    Check 'locked state file rolls back registry changes on failed commit' {
        $before = Snapshot
        $lock = [IO.File]::Open((Join-Path $root 'state.json'),[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        try { Reject { Invoke-XlDuckInstall -Action Install -Source $second @options } }
        finally { $lock.Dispose() }
        Assert ((Snapshot) -ceq $before) 'Failed installation left changed startup settings.'
    }
    Check 'unknown registered XlDuck is not overwritten' {
        $key.SetValue('OPEN4','/R "C:\Manual\XlDuck-AddIn64.xll"'); $before = Snapshot
        try { Reject { Invoke-XlDuckInstall -Action Install -Source $second @options } 'outside this installer'; Assert ((Snapshot) -ceq $before) 'Unmanaged registration changed.' }
        finally { $key.DeleteValue('OPEN4') }
    }
    Check 'corrupted payload fails before changing installation state' {
        $damaged = New-Package 'damaged'; [IO.File]::WriteAllText((Join-Path $damaged 'XlDuck.dll'),'corrupt')
        $before = Snapshot
        Reject { Invoke-XlDuckInstall -Action Install -Source $damaged @options } 'mismatch'
        Assert ((Snapshot) -ceq $before) 'Damaged package changed registration.'
    }
    Check 'download marking is not silently removed' {
        $downloaded = New-Package 'downloaded'
        Set-Content -LiteralPath (Join-Path $downloaded 'XlDuck-AddIn64.xll') -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3"
        $before = Snapshot
        Reject { Invoke-XlDuckInstall -Action Install -Source $downloaded @options } 'marked as downloaded'
        Assert ((Snapshot) -ceq $before) 'Downloaded payload changed registration.'
        Assert ((Get-Content -LiteralPath (Join-Path $downloaded 'XlDuck-AddIn64.xll') -Stream Zone.Identifier) -contains 'ZoneId=3') 'Zone marking removed.'
    }
    Check 'uninstall preserves newly added third-party entries and user data' {
        $key.SetValue('OPEN3','/R "C:\Other\new.xll"')
        [IO.File]::WriteAllText((Join-Path $root 'user-notes.txt'),'keep')
        $removed = Invoke-XlDuckInstall -Action Uninstall @options
        $entries = @(Get-XlDuckStartupEntries $key)
        Assert ($entries.Count -eq 3 -and $entries[2].name -eq 'OPEN2' -and $entries[2].value -eq '/R "C:\Other\new.xll"') 'Uninstall damaged new entry or left a gap.'
        Assert ((Get-Content -LiteralPath (Join-Path $root 'user-notes.txt')) -eq 'keep') 'Uninstall removed user file.'
        Assert (-not $removed.state.active) 'Uninstall left an active version.'
        Assert (Test-Path -LiteralPath (Join-Path $installed.directory 'XlDuck.dll')) 'Recovery copy removed.'
    }
    Check 'already-uninstalled setup is safe to repeat' {
        $before = Snapshot; $null = Invoke-XlDuckInstall -Action Uninstall @options
        Assert ((Snapshot) -ceq $before) 'Repeated uninstall changed third-party entries.'
    }
    Check 'malformed state cannot redirect installation paths' {
        $badRoot = Join-Path $testRoot 'bad-state'; [void][IO.Directory]::CreateDirectory($badRoot)
        [IO.File]::WriteAllText((Join-Path $badRoot 'state.json'),'{"schema":1,"product":"XlDuck","versions":["../../outside"],"active":"../../outside","previous":null}')
        Reject { Get-XlDuckInstallState $badRoot } 'Invalid installed version'
    }
    Check 'bootstrap ZIP contains only the actual minimal skill' {
        Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead((Join-Path $PackageDirectory 'agent/xlduck-skill.zip'))
        try {
            Assert ($zip.Entries.Count -eq 1 -and $zip.Entries[0].FullName -ceq 'xlduck/SKILL.md') 'Skill ZIP contains stale helpers or references.'
            $reader = [IO.StreamReader]::new($zip.Entries[0].Open())
            try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
            Assert ($content -ceq [IO.File]::ReadAllText((Join-Path $PackageDirectory 'agent/xlduck/SKILL.md'))) 'Skill ZIP differs from installed source.'
            Assert ($content.Length -lt 1500) 'Bootstrap contains too much implementation guidance.'
        } finally { $zip.Dispose() }
    }
    Check 'local skill installs, repeats safely and preserves a customized skill' {
        $skillSource = Join-Path $PackageDirectory 'agent/xlduck'; $skillTarget = Join-Path $testRoot 'skills/xlduck'
        Install-XlDuckLocalSkill $skillSource $skillTarget
        Install-XlDuckLocalSkill $skillSource $skillTarget
        [IO.File]::WriteAllText((Join-Path $skillTarget 'SKILL.md'),'user customization')
        Reject { Install-XlDuckLocalSkill $skillSource $skillTarget } 'will not overwrite'
        Assert ((Get-Content -LiteralPath (Join-Path $skillTarget 'SKILL.md')) -eq 'user customization') 'Customized skill was overwritten.'
    }
    Check 'launchers isolate inherited module paths, handle spaces and preserve exit codes' {
        $launcherRoot = Join-Path $testRoot 'launcher probes'
        $probeScripts = Join-Path $launcherRoot 'scripts'
        $foreignModules = Join-Path $testRoot 'foreign-modules'
        $foreignUtility = Join-Path $foreignModules 'Microsoft.PowerShell.Utility'
        [void][IO.Directory]::CreateDirectory($probeScripts)
        [void][IO.Directory]::CreateDirectory($foreignUtility)
        [IO.File]::WriteAllText((Join-Path $foreignUtility 'Microsoft.PowerShell.Utility.psd1'), "@{RootModule='bad.psm1';ModuleVersion='99.0';FunctionsToExport=@('Get-FileHash')}")
        [IO.File]::WriteAllText((Join-Path $foreignUtility 'bad.psm1'), "throw 'An incompatible module path was inherited.'")
        [IO.File]::WriteAllText((Join-Path $probeScripts 'Setup.ps1'), @'
param([string]$Action)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop' -or -not [Environment]::Is64BitProcess) { exit 91 }
$hash = (Get-FileHash -LiteralPath $PSCommandPath).Hash
if ($hash.Length -ne 64) { exit 92 }
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'observed-action.txt'),$Action)
exit 23
'@)
        $launchers = @{'Install.cmd'='Install';'Rollback.cmd'='Rollback';'Uninstall.cmd'='Uninstall';'Install Codex Skill.cmd'='InstallCodexSkill'}
        foreach ($name in $launchers.Keys) {
            $launcher = Join-Path $launcherRoot $name
            Copy-Item -LiteralPath (Join-Path $PackageDirectory $name) -Destination $launcher
            $start = [Diagnostics.ProcessStartInfo]::new($env:ComSpec,('/d /s /c ""' + $launcher + '" < nul"'))
            $start.UseShellExecute = $false; $start.CreateNoWindow = $true
            $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
            $start.EnvironmentVariables['PSModulePath'] = $foreignModules + ';' + (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/Modules')
            $probe = [Diagnostics.Process]::Start($start)
            try {
                $stdout = $probe.StandardOutput.ReadToEndAsync(); $stderr = $probe.StandardError.ReadToEndAsync()
                Assert ($probe.WaitForExit(30000)) 'Launcher probe did not exit.'
                Assert ($probe.ExitCode -eq 23) ("$name failed: " + $stderr.Result + $stdout.Result)
                Assert ([IO.File]::ReadAllText((Join-Path $probeScripts 'observed-action.txt')) -ceq $launchers[$name]) 'Launcher passed the wrong action.'
            } finally { $probe.Dispose() }
        }
    }
} finally {
    $key.Dispose()
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($testKey, $false)
    $allowed = [IO.Path]::GetFullPath((Join-Path $repo 'artifacts/setup-tests')).TrimEnd('\') + '\'
    if (-not [IO.Path]::GetFullPath($testRoot).StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe setup test cleanup.' }
    Assert-XlDuckPlainPath $testRoot
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
Write-Host "$passed passed, $failed failed."
if ($failed) { exit 1 }
