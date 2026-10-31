# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([ValidateSet('Check','Install','Rollback','Uninstall','InstallCodexSkill')][string]$Action = 'Check')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Setup.Core.ps1')
$package = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$installRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'XlDuck/install'

function Read-RegistryValue($Hive, [string]$Path, [string]$Name) {
    $key = $Hive.OpenSubKey($Path)
    try { if ($key) { return $key.GetValue($Name) } } finally { if ($key) { $key.Dispose() } }
}
function Get-PeMachine([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5a4d) { throw 'Not a Windows executable.' }
        $stream.Position = 0x3c; $offset = $reader.ReadInt32()
        $stream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x4550) { throw 'Invalid Windows executable.' }
        return $reader.ReadUInt16()
    } finally { $reader.Dispose() }
}
function Get-SetupIssues {
    if (-not [Environment]::Is64BitOperatingSystem -or -not [Environment]::Is64BitProcess -or
        $env:PROCESSOR_ARCHITECTURE -ne 'AMD64') { 'Use Windows x64 and 64-bit PowerShell; ARM and 32-bit Excel are not supported.' }
    $excelPath = Read-RegistryValue ([Microsoft.Win32.Registry]::CurrentUser) 'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' ''
    if (-not $excelPath) { $excelPath = Read-RegistryValue ([Microsoft.Win32.Registry]::LocalMachine) 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' '' }
    if (-not $excelPath -or -not [IO.File]::Exists($excelPath)) { 'Desktop Excel was not found. Install and open 64-bit Microsoft 365 / Excel 2021 or 2024 first.' }
    else {
        if ((Get-PeMachine $excelPath) -ne 0x8664) { 'Excel is not x64. This package cannot load into 32-bit or ARM Excel.' }
        if ([Diagnostics.FileVersionInfo]::GetVersionInfo($excelPath).FileMajorPart -ne 16) { 'Setup supports Office 16.x (Microsoft 365, Excel 2021/2024). See docs/INSTALL.md for manual registration.' }
    }
    $dotnetLocation = Read-RegistryValue ([Microsoft.Win32.Registry]::LocalMachine) 'SOFTWARE\dotnet\Setup\InstalledVersions\x64' 'InstallLocation'
    if (-not $dotnetLocation) { $dotnetLocation = Join-Path $env:ProgramFiles 'dotnet' }
    $dotnet = Join-Path $dotnetLocation 'dotnet.exe'
    $desktop = $false
    if ([IO.File]::Exists($dotnet) -and (Get-PeMachine $dotnet) -eq 0x8664) {
        $lines = @(& $dotnet --list-runtimes)
        $desktop = $LASTEXITCODE -eq 0 -and @($lines | Where-Object { $_ -match '^Microsoft.WindowsDesktop.App 10\.0\.\d+ \[' }).Count -gt 0
    }
    if (-not $desktop) { 'Install the .NET 10 Desktop Runtime (Windows x64): https://dotnet.microsoft.com/en-us/download/dotnet/10.0' }
    $webview = $false
    foreach ($hive in @([Microsoft.Win32.Registry]::CurrentUser,[Microsoft.Win32.Registry]::LocalMachine)) {
        $keyPath = 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
        if ($hive -eq [Microsoft.Win32.Registry]::LocalMachine) { $keyPath = $keyPath.Replace('SOFTWARE\','SOFTWARE\WOW6432Node\') }
        $version = Read-RegistryValue $hive $keyPath 'pv'
        $parsed = $null
        if ([version]::TryParse([string]$version,[ref]$parsed) -and $parsed -gt [version]'0.0.0.0') { $webview = $true }
    }
    if (-not $webview) { 'Install the WebView2 Evergreen Runtime: https://developer.microsoft.com/en-us/microsoft-edge/webview2/' }
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { 'Save your work and close Excel before installing, upgrading, rolling back or uninstalling. Setup never closes it.' }
}

$mutex = $null; $locked = $false
try {
    if (-not [Environment]::Is64BitProcess) { throw 'Run setup in 64-bit Windows PowerShell.' }
    if ($Action -eq 'Check' -or $Action -eq 'Install' -or $Action -eq 'InstallCodexSkill') {
        & (Join-Path $PSScriptRoot 'Verify-Package.ps1') -Directory $package
    }
    if ($Action -eq 'Check' -or $Action -eq 'Install') {
        $issues = @(Get-SetupIssues)
        foreach ($issue in $issues) { Write-Host "ACTION NEEDED: $issue" -ForegroundColor Yellow }
        if ($issues.Count) { throw 'Prerequisites are incomplete; no installation or Excel settings were changed.' }
        Write-Host 'PASS Windows x64, Excel x64, .NET 10 Desktop Runtime, WebView2, and Excel closed.'
        if ($Action -eq 'Check') { return }
    }
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $mutex = [Threading.Mutex]::new($false, ('Local\XlDuckSetup-' + $sid))
    try { $locked = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $locked = $true }
    if (-not $locked) { throw 'Another XlDuck setup is running. Finish it first.' }
    if ($Action -eq 'InstallCodexSkill') {
        $skillTarget = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.agents/skills/xlduck'
        Install-XlDuckLocalSkill -Source (Join-Path $package 'agent/xlduck') -Destination $skillTarget
        Write-Host "Installed bootstrap skill for local Codex: $skillTarget"
        Write-Host 'Start a new Codex chat and select the xlduck skill. Excel-sidebar skills are installed separately; see START-HERE.html.'
        return
    }
    $result = Invoke-XlDuckInstall -Action $Action -Source $package -Root $installRoot -OptionsKey 'Software\Microsoft\Office\16.0\Excel\Options'
    if ($Action -eq 'Uninstall') {
        Write-Host 'XlDuck startup registration removed. Connection settings, credentials and workbooks are unchanged.'
        Write-Host "Recovery copies remain at $installRoot. You may delete that installation folder after checking it for your own files."
        Write-Host 'Remove any separately installed agent skill through its provider. Local Codex skill: %USERPROFILE%\.agents\skills\xlduck'
    } else {
        Write-Host "XlDuck registered for normal Excel startup: $($result.directory)"
        Write-Host 'Open Excel normally and test =DuckVersion() and =DuckQueryOutScalar("SELECT 42").'
        Write-Host 'For optional agent setup, open START-HERE.html. Use Rollback.cmd to return to the preceding installation.'
    }
} catch { Write-Host "SETUP STOPPED: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }
finally { if ($locked) { $mutex.ReleaseMutex() }; if ($mutex) { $mutex.Dispose() } }
