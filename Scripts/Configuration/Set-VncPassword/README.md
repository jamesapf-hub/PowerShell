# Set VNC Password Configuration


## Overview

Configures and resets TightVNC / UltraVNC server administrative passwords in the system registry.


> [!NOTE]
> **Log File Location:** `C:\Logs\Set-VncPassword\VncPasswordReset.txt` (or `$env:SystemDrive\Logs\Set-VncPassword\VncPasswordReset.txt`)

## Prerequisites

**OS Support:** Windows 10 / 11 / Windows Server
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console (passing `-Password` or setting `$env:VNC_PASSWORD` / `$env:DATTO_VNC_PASSWORD`):
   ```powershell
   .\Set-VncPassword.ps1 -Password "YourVncPassword"
   ```

## Fast Execute
> [!TIP]
> **Short Branded URL (phnx.it):**
> ```powershell
> $env:VNC_PASSWORD="YourVncPassword"; irm https://phnx.it/VNCPASS | iex
> ```
>
> **Full Transparent GitHub Command:**
> ```powershell
> $env:VNC_PASSWORD="YourVncPassword"; iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Configuration/Set-VncPassword/Set-VncPassword.ps1")
> ```
