@echo off
setlocal
rem Use Windows PowerShell's own modules even when launched from PowerShell 7.
set "PSModulePath=%SystemRoot%\System32\WindowsPowerShell\v1.0\Modules"
set "XLDUCK_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "XLDUCK_PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%XLDUCK_PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Setup.ps1" -Action Uninstall
set "XLDUCK_EXIT=%ERRORLEVEL%"
pause
exit /b %XLDUCK_EXIT%
