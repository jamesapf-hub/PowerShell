@echo off
title Microsoft 365 Shared Mailbox Access Viewer
cd /d "%~dp0"
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0SharedMailboxAccessViewer.ps1"
exit /b 0