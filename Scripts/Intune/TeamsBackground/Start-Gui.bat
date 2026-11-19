@echo off
title Microsoft Teams Background Win32 App Packager for Intune
echo ============================================================
echo Launching Teams Background Packager GUI...
echo ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0TeamsBackgroundPackager.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Packager closed with exit code %ERRORLEVEL%. Press any key to exit.
    pause
)
