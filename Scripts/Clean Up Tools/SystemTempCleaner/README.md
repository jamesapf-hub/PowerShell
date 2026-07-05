# System Temp Files & Recycle Bin Cleaner Guide

---

## Overview

This PowerShell script cleans Windows system temporary files, user temp files, Windows Prefetch cache, and empties the Recycle Bin. It operates in a dry-run (WhatIf) mode first by default to allow safe review before any files are modified.

### Key Features
* **System Temp Purge:** Deletes files in `C:\Windows\Temp` recursively.
* **User Temp Purge:** Deletes files in the active user's AppData Temp folder (`$env:TEMP`).
* **Prefetch Cleanup:** Purges the Windows Prefetch cache directory (`C:\Windows\Prefetch`).
* **Recycle Bin Empty:** Programmatically empties the system Recycle Bin.
* **Interactive Prompts:** Supports dry-runs and safe confirmation prompts before execution.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required (elevation check included)
* **Execution Policy:** RemoteSigned or Bypass

---

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console with administrative privileges (Run as Administrator).
2. Execute the script to initiate the WhatIf dry-run check. It will print out each folder it targets.
3. A prompt will ask: `WhatIf dry-run completed. Do you want to run this for real now? (Y/N)`
4. Type `Y` or `Yes` and press Enter to execute the real cleanup.
5. Alternatively, run the script with the `-Force` switch to bypass all prompts and immediately execute the cleanup.

### 2. Logging & Outputs
* Status logs for directories scanned and deleted are printed directly to the console.

---

## Command

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\01_SystemTempCleaner.ps1
```
