# Install Windows App (LTSC IoT Provisioning)


## Overview

Performs offline provisioning of the Microsoft Windows App for Windows 10 IoT LTSC environments. It pulls the latest verified MSIX/MSIXBUNDLE payload from the Microsoft CDN repository and registers it machine-wide.


> [!NOTE]
> **Log File Location:** `C:\Logs\Install-WindowsAppLTSC\ltsc_app_update.log` (or `$env:SystemDrive\Logs\Install-WindowsAppLTSC\ltsc_app_update.log`)

## Prerequisites

**OS Support:** Windows 10 / 11 / Windows Server / LTSC IoT
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Install-WindowsAppLTSC.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Kiosk/Install-WindowsAppLTSC/Install-WindowsAppLTSC.ps1")
> ```
