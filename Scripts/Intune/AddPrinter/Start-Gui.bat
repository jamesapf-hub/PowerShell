@echo off
title Intune Printer Packager Launcher
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "AddPrinter.ps1"
pause
