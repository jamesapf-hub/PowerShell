# Windows Upgrade Folder Cleanup Guide

## Overview
This script checks for the presence of leftover Windows upgrade cache folders (`C:\$WINDOWS.~WS` and `C:\ESD`), takes ownership and grants permissions recursively, and deletes them to recover disk space.

### Key Features
* **Targeted Folders:** Targets `C:\$WINDOWS.~WS` (installation workspace) and `C:\ESD` (upgrade ESD files).
* **Automatic Ownership:** Natively runs `takeown` and `icacls` commands to grant full access to administrators before deletion.
* **Dry-Run Mode (WhatIf):** Runs a WhatIf dry-run first to check folders before applying changes.

> [!NOTE]
> **Log File Location:** `C:\Logs\TidyUpWindows11\TidyUpWindows11_DDMMYY.log` (or `$env:SystemDrive\Logs\TidyUpWindows11\TidyUpWindows11_DDMMYY.log`)

## Prerequisites
OS Support: Windows 10 / 11
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required (elevation check included)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console with administrative privileges (Run as Administrator).
2. Run the script to perform a dry-run check of the folders.
3. A prompt will ask: `WhatIf dry-run completed. Do you want to run this for real now? (Y/N)`
4. Type `Y` or `Yes` and press Enter to take ownership and recursively delete the folders.
5. Alternatively, run with the `-Force` switch to immediately clean the folders without prompts.

### 2. Logging & Outputs
* Status logs for folder detection, ownership acquisition, and deletion are printed directly to the console.


## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean Up Tools/TidyUpWindows11/06_TidyUpWindows11.ps1")
> ```
