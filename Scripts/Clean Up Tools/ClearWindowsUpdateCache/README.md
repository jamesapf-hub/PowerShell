# Windows Update Cache Cleanup Guide

## Overview
This PowerShell script stops Windows Update services (`wuauserv` and `bits`), purges the `C:\Windows\SoftwareDistribution\Download` cache directory, and restarts the services. It runs in a dry-run (WhatIf) mode first to show potential cleanup statistics, allowing administrators to confirm execution before applying changes.

### Key Features
* **Dry-Run Mode (WhatIf):** Scans and displays the count and size of cached update files before performing any actions.
* **Service Management:** Handles stopping and restarting of dependent Windows Update services safely.
* **Robust Fallback:** If native file deletion fails, it automatically falls back to a robocopy mirror sync to purge locked directories.
* **Unified Logging:** Appends timestamped execution logs directly to `$env:SystemDrive\Logs\ClearWindowsUpdateCache`.

> [!NOTE]
> **Log File Location:** C:\Logs\ClearWindowsUpdateCache\SD_Clear_DDMMYY.log (or $env:SystemDrive\Logs\ClearWindowsUpdateCache\SD_Clear_DDMMYY.log)

## Prerequisites
OS Support: Windows 10 / 11 / Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required (Administrator elevation check included)
Execution Policy: RemoteSigned or Bypass

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console with administrative privileges (Run as Administrator).
2. Execute the script to initiate the standard dry-run. It will scan `C:\Windows\SoftwareDistribution\Download` and output how many files and MB can be cleaned.
3. A prompt will ask: `WhatIf dry-run completed. Do you want to run this for real now? (Y/N)`
4. Type `Y` or `Yes` and press Enter to execute the real cleanup, stopping the update services, performing the purge, and restarting services.
5. Alternatively, run the script with the `-Force` switch to bypass all prompts and immediately execute the cleanup.

### 2. Logging & Outputs
* Standard outputs and service status logs are written to the console in real-time.
* A persistent execution log is saved under `$env:SystemDrive\Logs\ClearWindowsUpdateCache\SD_Clear_DDMMYY.log` using the UK date format (DDMMYY).

## Local Execution Command

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\04_ClearWindowsUpdateCache.ps1
```

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/ClearWindowsUpdateCache/04_ClearWindowsUpdateCache.ps1")
> ```
