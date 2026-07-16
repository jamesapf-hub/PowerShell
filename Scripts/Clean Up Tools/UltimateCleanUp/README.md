# Ultimate PC CleanUp Utility

## Overview
This is the self-contained package for the **Ultimate PC CleanUp Utility**. It includes the central dark-themed WPF dashboard/CLI runner (`Start-UltimateCleanUp.ps1`), the double-click launcher (`Run-UltimateCleanUp.bat`), and all 6 sub-cleanup scripts flatly packaged inside the `Scripts/` directory.

### Discovered Tasks Included:
1. **System Temp Files & Recycle Bin Cleaner:** Purges system-level temporary files, Prefetch cache, and empties the Recycle Bin.
2. **Browser Cache Cleaner:** Simulates Google Chrome and Microsoft Edge cache clearing.
3. **Adobe Orphan Installer Patch Cleaner:** Scans and removes orphaned Adobe update patch files (`.msp`) using COM APIs.
4. **Clear Windows Update Cache:** Stops Windows Update services and purges the `SoftwareDistribution` cache folder.
5. **Orphaned AD Domain User Profile Cleaner:** Identifies and cleanly deletes Active Directory profiles of users who no longer exist.
6. **Windows Upgrade Folder Cleanup:** Takes ownership and deletes leftover Windows upgrade cache folders (`C:\$WINDOWS.~WS` and `C:\ESD`).

> [!NOTE]
> **Windows Update Cache Log Location:** `C:\Logs\UltimateCleanUp\SD_Clear_DDMMYY.log` (or `$env:SystemDrive\Logs\UltimateCleanUp\SD_Clear_DDMMYY.log`)

## Prerequisites
OS Support: Windows 10 / 11 / Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required (the batch launcher handles UAC elevation prompts automatically).

## How to Download & Run

### 1. Download the Package
Simply download this entire `UltimateCleanUp` directory from the repository.

### 2. Run the Dashboard (GUI)
Double-click **`Run-UltimateCleanUp.bat`** inside this folder.
* Note: The script handles UAC elevation checks automatically. If it requires administrator permissions, a prompt will appear, and it will relaunch in a new elevated window.

### 3. Run via Command-Line (CLI)
Open PowerShell as Administrator inside this folder and run:
* **List discovered tasks:**
  ```powershell
  .\Start-UltimateCleanUp.ps1 -ListTasks
  ```
* **Run all tasks silently:**
  ```powershell
  .\Start-UltimateCleanUp.ps1 -RunAll
  ```
* **Run specific tasks by name:**
  ```powershell
  .\Start-UltimateCleanUp.ps1 -RunTasks "SystemTempCleaner, ClearWindowsUpdateCache"
  ```

## Folder Structure

```text
UltimateCleanUp/
├── Run-UltimateCleanUp.bat    # Double-click launcher
├── Start-UltimateCleanUp.ps1  # Main runner dashboard script
├── README.md                  # This documentation guide
└── Scripts/                   # Sub-cleanup scripts folder
    ├── 01_SystemTempCleaner.ps1
    ├── 02_BrowserCacheCleaner.ps1
    ├── 03_CleanAdobeOrphanInstallers.ps1
    ├── 04_ClearWindowsUpdateCache.ps1
    ├── 05_RemoveOldUserProfiles.ps1
    └── 06_TidyUpWindows11.ps1
```


## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean Up Tools/UltimateCleanUp/Start-UltimateCleanUp.ps1")
> ```
