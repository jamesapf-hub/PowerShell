@echo off
title M365 Admin Console Launcher
cd /d "%~dp0"

:: Check standard installation paths first (in case PATH isn't updated yet)
set "PWSH_PATH="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" (
    "%ProgramFiles%\PowerShell\7\pwsh.exe" -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=%ProgramFiles%\PowerShell\7\pwsh.exe"
)
if not defined PWSH_PATH if exist "%LocalAppData%\Microsoft\powershell\pwsh.exe" (
    "%LocalAppData%\Microsoft\powershell\pwsh.exe" -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=%LocalAppData%\Microsoft\powershell\pwsh.exe"
)
if not defined PWSH_PATH if exist "%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe" (
    "%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe" -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe"
)
if not defined PWSH_PATH if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" (
    "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
)

:: If not found in standard paths, check if pwsh is in PATH and actually works.
:: Running 'pwsh -Command' prevents false positives from Windows Store App Execution Aliases.
if not defined PWSH_PATH (
    pwsh -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=pwsh"
)

:: Launch if found and working
if defined PWSH_PATH (
    echo Launching with PowerShell 7...
    "%PWSH_PATH%" -NoProfile -ExecutionPolicy Bypass -File "Console.ps1"
    goto end
)

echo ======================================================================
echo       ERROR: PowerShell 7 is required but was not found.
echo ======================================================================
echo.
echo The M365 Admin Console requires PowerShell 7.2 or higher.
echo Your system is currently running legacy Windows PowerShell 5.1.
echo.
set /p choice="Would you like to install the latest PowerShell 7 now? (Y/N) [Y]: "
if /i "%choice%"=="N" (
    echo Launch cancelled.
    goto end
)

echo.
echo Attempting to install PowerShell 7...
echo.

:: 1. Try winget first (Standard on Windows 10/11)
where winget >nul 2>nul
if %ERRORLEVEL% equ 0 (
    echo [1/2] Installing via winget...
    winget install --id Microsoft.PowerShell --source winget
    if %ERRORLEVEL% equ 0 goto success
)

:: 2. Fallback to Microsoft's official installer script
echo [2/2] winget not found or failed. Installing via official script...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $t = Join-Path $env:TEMP 'install-pwsh.ps1'; Invoke-RestMethod 'https://aka.ms/install-powershell.ps1' -OutFile $t; & $t; Remove-Item $t -ErrorAction SilentlyContinue"

:success
echo.
echo ======================================================================
echo   PowerShell 7 installation process completed.
echo   Configuring environment and preparing to launch console...
echo ======================================================================
echo.
timeout /t 3 >nul

:: Re-evaluate PWSH_PATH
set "PWSH_PATH="
if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PWSH_PATH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PWSH_PATH if exist "%LocalAppData%\Microsoft\powershell\pwsh.exe" set "PWSH_PATH=%LocalAppData%\Microsoft\powershell\pwsh.exe"
if not defined PWSH_PATH if exist "%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe" set "PWSH_PATH=%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe"
if not defined PWSH_PATH if exist "%ProgramFiles(x86)%\PowerShell\7\pwsh.exe" set "PWSH_PATH=%ProgramFiles(x86)%\PowerShell\7\pwsh.exe"
if not defined PWSH_PATH (
    pwsh -Command "exit 0" >nul 2>nul
    if not errorlevel 1 set "PWSH_PATH=pwsh"
)

if defined PWSH_PATH (
    echo Launching with PowerShell 7...
    "%PWSH_PATH%" -NoProfile -ExecutionPolicy Bypass -File "Console.ps1"
) else (
    echo [WARNING] PowerShell 7 installation completed but the executable could not be located.
    echo Please try to open LaunchConsole.bat again.
    pause
)

:end
echo.
echo Process finished. Press any key to close launcher window...
pause
