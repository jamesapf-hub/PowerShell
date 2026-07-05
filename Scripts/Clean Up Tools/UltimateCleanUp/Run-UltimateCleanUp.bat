@echo off
:: Launches the Ultimate PC CleanUp Utility PowerShell GUI
:: The PowerShell script itself handles Administrator check and auto-elevates if needed.

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Start-UltimateCleanUp.ps1" %*
