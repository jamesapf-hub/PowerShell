@echo off
title Edge Profile Shortcut Builder
echo ============================================================
echo Launching Edge Profile Shortcut Builder GUI...
echo ============================================================
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0EdgeProfileBuilder.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script closed with exit code %ERRORLEVEL%. Press any key to exit.
    pause
)
