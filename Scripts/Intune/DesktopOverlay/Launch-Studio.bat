@echo off
title IT Support Overlay Studio
echo ============================================================
echo Opening IT Support Overlay Studio and Package Builder...
echo ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0OverlayStudio.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script encountered an error. Press any key to exit.
    pause
)
