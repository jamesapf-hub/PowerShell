# Schedule Kiosk App Runtime Fix


## Overview

Creates a scheduled task to run Kiosk Shell App runtime fixes on user authentication to prevent logon UI hanging states.


> [!NOTE]
> **Log File Location:** `C:\Logs\Schedule-KioskAppRuntimeFix\Kiosk_Runtime_Fix.txt` (or `$env:SystemDrive\Logs\Schedule-KioskAppRuntimeFix\Kiosk_Runtime_Fix.txt`)

## Prerequisites

**OS Support:** Windows 10 / 11
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Schedule-KioskAppRuntimeFix.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Kiosk/Schedule-KioskAppRuntimeFix/Schedule-KioskAppRuntimeFix.ps1")
> ```
