# Install Windows Updates (LTSC)


## Overview

Triggers Windows Update check and installation using the native Windows Update Agent COM objects. Operates without any external module dependencies.


> [!NOTE]
> **Log File Location:** `C:\Logs\Install-WindowsUpdatesLTSC\WindowsUpdate_DDMMYY.log` (or `$env:SystemDrive\Logs\Install-WindowsUpdatesLTSC\WindowsUpdate_DDMMYY.log`)

## Prerequisites

**OS Support:** Windows 10 / 11 / Windows Server / LTSC IoT
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Install-WindowsUpdatesLTSC.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Software/Install-WindowsUpdatesLTSC/Install-WindowsUpdatesLTSC.ps1")
> ```
