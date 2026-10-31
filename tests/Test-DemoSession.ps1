# Copyright (c) 2026 James Thompson
# SPDX-License-Identifier: MPL-2.0
# Reproduce a null COM Open result without starting or attaching to Excel.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../scripts/DemoSession.ps1')
$requestedPath = (Resolve-Path (Join-Path $PSScriptRoot '../outputs/01a16b0a-0074-7383-ba33-19bcf4c27b0f/XlDuck Research Demo.xlsx')).Path
$requestedBook = [pscustomobject]@{Name=[IO.Path]::GetFileName($requestedPath);FullName=$requestedPath}
$decoyBook = [pscustomobject]@{Name=$requestedBook.Name;FullName=(Join-Path $env:TEMP $requestedBook.Name)}
$blankBook = [pscustomobject]@{Name='Owned blank';FullName='Owned blank';Closed=$false}
$blankBook | Add-Member ScriptMethod Close { param($save) $this.Closed = $true }
$mockBooks = [Collections.ArrayList]@($decoyBook,$requestedBook,$blankBook)
Add-Member -InputObject $mockBooks -MemberType ScriptMethod -Name Open -Value { param($path) return $null }
$mockApplication = [pscustomobject]@{Workbooks=$mockBooks;Calculation=-4105}
$mockApplication | Add-Member ScriptMethod Run { param($macro,$argument) }
$mockApplication | Add-Member ScriptMethod RegisterXLL { param($path) return $true }
function Copy-XlDuckDevelopmentBuild { param($AddInPath) return $AddInPath }
function Assert-XlDuckExcelSession { param($Session) }
function New-XlDuckExcelSession { param($AddInPath) return [pscustomobject]@{Application=$mockApplication;Workbook=$blankBook;WorkbookName=$blankBook.Name} }
function Close-XlDuckExcelSession { param($Session) throw 'Unexpected failure cleanup in successful opening regression.' }
# The only COM-specific operation remaining is releasing the known blank RCW.
# Exercise the real opening logic through a function copy with that cleanup removed
# because these objects are deliberately ordinary objects rather than Excel RCWs.
$demoBody = (Get-Command New-XlDuckDemoSession).ScriptBlock.ToString()
$scriptsRoot = (Resolve-Path (Join-Path $PSScriptRoot '../scripts')).Path
$demoBody = $demoBody.Replace('$PSScriptRoot', ("'" + $scriptsRoot.Replace("'", "''") + "'"))
$demoBody = $demoBody.Replace('[void][Runtime.InteropServices.Marshal]::ReleaseComObject($blank)', '$null = $blank')
Set-Item Function:New-XlDuckDemoSession -Value ([scriptblock]::Create($demoBody))
$openedSession = New-XlDuckDemoSession -WorkbookPath $requestedPath -AddInPath 'mock.xll'
if (-not [object]::ReferenceEquals($openedSession.Workbook, $requestedBook)) { throw 'Did not recover the requested workbook by its complete path.' }
if (-not $blankBook.Closed) { throw 'Did not close the owned blank after resolving the requested workbook.' }
if ($openedSession.WorkbookName -ne $requestedBook.Name) { throw 'Did not retain the requested workbook identity.' }
Write-Host 'PASS null Open result resolves the exact requested path in the owned instance, ignoring another book with the same filename.'
