# Scheduled Task Windows App Log Off Watchdog


## Overview

Sets up a watchdog scheduled task to monitor UWP app processes. If the primary kiosk app closes, the script automatically logs off the local user to keep the terminal secure.


> [!NOTE]
> **Log File Location:** `C:\Logs\ScheduledTaskWindowsAppLogOff\KioskWatchdog.log` (or `$env:SystemDrive\Logs\ScheduledTaskWindowsAppLogOff\KioskWatchdog.log`)

## Prerequisites

**OS Support:** Windows 10 / 11
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\ScheduledTaskWindowsAppLogOff.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Kiosk/ScheduledTaskWindowsAppLogOff/ScheduledTaskWindowsAppLogOff.ps1")
> ```
