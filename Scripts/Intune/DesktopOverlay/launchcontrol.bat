@echo off
title IT Support Overlay Local Launcher
echo ============================================================
echo Launching IT Support Overlay (Local Test Mode)
echo Press Ctrl+C or kill PowerShell process to close overlay
echo ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Start-ITOverlay.ps1"
pause
