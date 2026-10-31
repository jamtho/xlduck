# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
param([string]$AddInPath, [string]$RuntimeRoot, [string]$FrameworkProbe)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/ExcelSession.ps1')
if (-not $AddInPath) { $AddInPath = Join-Path $PSScriptRoot '../XlDuck/bin/Debug/net8.0-windows/XlDuck-AddIn64.xll' }
Add-Type @'
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
public sealed class ProtectedViewReaderFixture : IDisposable {
    private readonly TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
    private Task loop;
    public int Requests;
    public volatile string Body = "n\n42\n";
    public string Url { get; private set; }
    public ProtectedViewReaderFixture() {
        listener.Start(); Url = "http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/test.csv";
        loop = Task.Run((Action)Serve);
    }
    private void Serve() {
        try { while (true) {
            using (var client = listener.AcceptTcpClient()) using (var stream = client.GetStream()) {
                client.ReceiveTimeout = 3000;
                var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, true);
                var request = reader.ReadLine(); string line;
                while (!string.IsNullOrEmpty(line = reader.ReadLine())) { }
                Interlocked.Increment(ref Requests);
                var body = Encoding.UTF8.GetBytes(Body);
                var header = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: text/csv\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n");
                stream.Write(header, 0, header.Length);
                if (request != null && !request.StartsWith("HEAD ")) stream.Write(body, 0, body.Length);
            }
        } } catch (SocketException) { } catch (ObjectDisposedException) { }
    }
    public void Dispose() { listener.Stop(); loop.Wait(4000); }
}
'@
$directory = Join-Path (Split-Path $PSScriptRoot -Parent) ('artifacts/protected-view-tests/' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($directory)
$path = Join-Path $directory 'Owned download fixture.xlsx'
$server = [ProtectedViewReaderFixture]::new()
$session = $null; $sheet = $null; $view = $null; $protectedBook = $null; $protectedSheet = $null
function Test-SandboxedWorkbook($Workbook) {
    $owner = $Workbook.Application
    try { return $owner.IsSandboxed }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($owner) }
}
try {
    $build = Copy-XlDuckDevelopmentBuild $AddInPath
    if ($RuntimeRoot) {
        $session = New-XlDuckExcelSession -AddInPath (Copy-XlDuckDevelopmentBuild $FrameworkProbe)
        $session.Application.Run('XlDuckProbeRuntimeRoot', [IO.Path]::GetFullPath($RuntimeRoot)) | Out-Null
        if (-not $session.Application.RegisterXLL($build)) { throw 'Add-in failed to load.' }
    } else { $session = New-XlDuckExcelSession -AddInPath $build }
    $excel = $session.Application; $sheet = $session.Workbook.Worksheets.Item(1)
    Invoke-XlDuckExcel { $sheet.Range('A1').Formula2 = '=DuckVersion()' }
    Invoke-XlDuckExcel { $sheet.Range('A2').Formula2 = '=DuckQueryOutScalar("SELECT n FROM read_csv(?)", B2)' }
    Invoke-XlDuckExcel { $sheet.Range('B2').Value2 = $server.Url }
    [void](Wait-XlDuckCell $sheet 'A2' -Until { param($v) $v -eq 42 })
    Invoke-XlDuckExcel { $session.Workbook.SaveAs($path, 51) }
    Invoke-XlDuckExcel { $session.Workbook.Close($false) }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $sheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($session.Workbook); $session.Workbook = $null
    $server.Body = "n`n99`n"; $before = $server.Requests
    # Mark only the workbook this test created. Do not alter Trust Center or any downloaded user file.
    Set-Content -LiteralPath $path -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding ASCII
    $view = Invoke-XlDuckExcel { $excel.ProtectedViewWindows.Open($path) }
    $protectedBook = $view.Workbook; $protectedSheet = $protectedBook.Worksheets.Item(1)
    if ($excel.Workbooks.Count -ne 0 -or -not (Test-SandboxedWorkbook $protectedBook)) { throw 'Fixture did not enter Protected View.' }
    Invoke-XlDuckExcel { $excel.CalculateFull() }
    # Allow host reconciliation and any accidental background work to reach the loopback server.
    Start-Sleep -Milliseconds 1500
    if ($protectedSheet.Range('A2').Value2 -ne 42 -or $server.Requests -ne $before) { throw 'A protected workbook ran a source query.' }
    # Excel-DNA may consume a macro exception instead of exposing a COM exception.
    # Assert the effects, not the exception transport.
    try { $excel.Run('DuckRestartWorkbook', $protectedBook.Name) | Out-Null } catch { }
    Start-Sleep -Milliseconds 1000
    if ($excel.Workbooks.Count -ne 0 -or $server.Requests -ne $before -or $protectedSheet.Range('A2').Value2 -ne 42) { throw 'A host command admitted a protected workbook.' }
    Write-Host 'PASS Protected View retains cached results and makes no external read'
    $ordinary = $null; $ordinarySheet = $null
    try {
        $ordinary = Invoke-XlDuckExcel { $excel.Workbooks.Add() }; $ordinarySheet = $ordinary.Worksheets.Item(1)
        Invoke-XlDuckExcel { $ordinarySheet.Range('A1').Formula2 = '=DuckQueryOutScalar("SELECT 7")' }
        [void](Wait-XlDuckCell $ordinarySheet 'A1' -Until { param($v) $v -eq 7 })
        if ((Test-SandboxedWorkbook $ordinary) -or -not (Test-SandboxedWorkbook $protectedBook) -or $server.Requests -ne $before) { throw 'Protected and ordinary workbook identities were confused.' }
    } finally {
        if ($ordinarySheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ordinarySheet) }
        if ($ordinary) { $ordinary.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ordinary) }
    }
    Write-Host 'PASS an ordinary workbook computes while a protected workbook remains dormant'
    # Exercise Enable Editing on our own known, generated fixture; no prompt or policy is bypassed.
    $session.Workbook = Invoke-XlDuckExcel { $view.Edit() }
    $session.WorkbookName = $session.Workbook.Name
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($protectedSheet); $protectedSheet = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($protectedBook); $protectedBook = $null
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view); $view = $null
    $sheet = $session.Workbook.Worksheets.Item(1)
    [void](Wait-XlDuckCell $sheet 'A2' -Until { param($v) $v -eq 99 })
    if ($server.Requests -le $before -or $excel.ProtectedViewWindows.Count -ne 0) { throw 'Enable Editing did not start a fresh ordinary session.' }
    Write-Host 'PASS Excel Enable Editing starts XlDuck automatically and reads fresh data'
} finally {
    foreach ($value in @($protectedSheet, $protectedBook, $sheet)) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } }
    if ($view) { $view.Close() | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
    if ($session) { Close-XlDuckExcelSession $session }
    $server.Dispose()
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $key = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($path).ToUpperInvariant()))).Replace('-', '') }
    finally { $sha.Dispose() }
    $record = Join-Path $env:LOCALAPPDATA ("XlDuck/connections/$key.json")
    if (Test-Path -LiteralPath $record) { Remove-Item -LiteralPath $record }
}
