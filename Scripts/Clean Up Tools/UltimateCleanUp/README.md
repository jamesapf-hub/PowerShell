# Ultimate Windows CleanUp Utility

---

## Overview

A premium, modular system cleaner optimized for enterprise Windows workstations and thin clients. It performs deep storage cleaning, purges obsolete local cache repositories, clears Windows Update files, and cleans up orphaned user profiles safely.

> [!NOTE]
> **Windows Update Cache Log Location:** `C:\Logs\UltimateCleanUp\SD_Clear_DDMMYY.log` (or `$env:SystemDrive\Logs\UltimateCleanUp\SD_Clear_DDMMYY.log`)

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.

---

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### Running the Main Cleanup Dashboard
1. Open Windows PowerShell as an Administrator.
2. Execute the main script loader:
   ```powershell
   .\Start-UltimateCleanUp.ps1
   ```
3. Select the modules to run (e.g., Disk Cleanup, Temp Files, Windows Update Cache, Old Profiles).
4. View the real-time execution log console on the right side of the screen.

#### Running Cache Cleanup Directly (Standalone)
To execute the Windows Update cache clearing module directly without the GUI console:
1. Open PowerShell as an Administrator.
2. Execute the standalone module:
   ```powershell
   .\Scripts\04_ClearWindowsUpdateCache.ps1
   ```
