@echo off
title Entra ID MFA SMS Deprecation Checker Launcher
echo ======================================================================
echo Launching Entra ID MFA SMS Deprecation Checker (PowerShell 7)
echo ======================================================================
echo.

where pwsh >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Located pwsh.exe in system PATH. Launching GUI...
    pwsh.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Start-MFACheckerGUI.ps1"
    goto :end
)

if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" (
    echo [INFO] Located PowerShell 7 at "%ProgramFiles%\PowerShell\7\pwsh.exe". Launching GUI...
    "%ProgramFiles%\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -NoProfile -File "%~dp0Start-MFACheckerGUI.ps1"
    goto :end
)

if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" (
    echo [INFO] Located PowerShell 7 at "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe". Launching GUI...
    "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -NoProfile -File "%~dp0Start-MFACheckerGUI.ps1"
    goto :end
)

echo [WARNING] PowerShell 7 (pwsh.exe) was not found in standard installation paths.
echo Recent Microsoft Graph modules perform best under PowerShell 7+.
echo.
echo You can install PowerShell 7 via command line:
echo   winget install --id Microsoft.PowerShell --source winget
echo.
echo Attempting fallback launch via Windows PowerShell 5.1...
echo.
pause
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Start-MFACheckerGUI.ps1"

:end
echo.
echo Launcher finished.
pause
