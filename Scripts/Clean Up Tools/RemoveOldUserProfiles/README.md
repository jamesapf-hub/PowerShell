# Remove Old User Profiles Guide

## Overview
> **Short Description:** Removes old user profiles from the system.

The **User Profile Cleanup Tool** (`05_RemoveOldUserProfiles.ps1`) is an automated PowerShell utility designed to detect and safely purge orphaned or deleted Active Directory user profiles from Windows machines. It intelligently detects live users to prevent accidental deletion, forcibly bypasses NTFS restrictions to report highly accurate folder sizes, and dynamically swaps between a graphical window and a text-based console menu depending on the execution environment (e.g. Desktop vs. Datto RMM).

### Key Features
* **[ACTIVE USER] Protection:** Deeply scans the `HKEY_USERS` registry hive and `Win32_ComputerSystem` WMI class to detect the live console user, tagging them and un-ticking them by default to prevent accidental deletion.
* **Safe Simulation (-WhatIf / -DryRun):** Safely simulate profile deletion without actually deleting any files or registry keys.
* **Automated Mode (-Force):** Completely bypasses all interactive menus for fully automated, silent cleanup deployments.
* **Fallback Console Mode (-ConsoleMode):** Forces a text-based selection menu rather than a graphical pop-up window, perfect for remote execution.
* **Ghost SID Purging:** Automatically detects completely corrupted SIDs missing their Active Directory bindings and registry folder paths, and forcefully purges them from the system.

> [!NOTE]
> **Log File Location:** `$env:SystemDrive\Logs\RemoveOldUserProfiles\CleanupLog_YYYYMMDD_HHMMSS.log` (Output dynamically written to console and file)

## Prerequisites
OS Support: Windows 10 / 11 and Windows Server
PowerShell: Windows PowerShell 5.1+
Permissions: Local Administrator rights required (Runs effectively as SYSTEM)
Execution Policy: Bypass or RemoteSigned
Dependencies: None (Natively utilizes Windows Forms, WMI, and FileSystemObject COM)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

**Interactive Execution (Standard)**
If you run the script directly from a standard PowerShell prompt, it will scan the machine, calculate folder sizes, and pop up a **Graphical Window**.
```powershell
.\05_RemoveOldUserProfiles.ps1
```
Simply un-tick any profiles you want to keep, and click **Delete Selected**. The active user will be automatically un-ticked for your protection.

**Remote Execution via Datto / Web Shell**
If you run the script through a background RMM like Datto, or force it using the `-ConsoleMode` switch, it will detect the restricted environment and launch the **Text-Based Menu**.
```powershell
.\05_RemoveOldUserProfiles.ps1 -ConsoleMode
```
You will see a numbered list of profiles (e.g., `[0]`, `[1]`, `[2]`). 
* Type specific numbers separated by commas to delete them (e.g., `0,2,3`).
* Type `all` to purge everything in the list.
* Type `none` to cancel and exit safely.

**Fully Automated Cleanup (No Prompts)**
If you are deploying this script as a scheduled job across hundreds of machines and want it to silently delete all orphaned profiles with zero human interaction:
```powershell
.\05_RemoveOldUserProfiles.ps1 -Force
```

**Safe Simulation (Dry Run)**
If you want to see exactly what the script *would* delete if you ran it on a fully automated schedule, without risking any actual data loss:
```powershell
.\05_RemoveOldUserProfiles.ps1 -DryRun
```

### 2. Logging & Outputs
The script provides deeply color-coded console output during its execution. When it finishes analyzing and calculating profile sizes, it will log all actions, including the forced purging of corrupted registry keys or stranded folders. Logs are also automatically dumped to a time-stamped text file in the root `C:\Logs` directory for auditing purposes.

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/RemoveOldUserProfiles/05_RemoveOldUserProfiles.ps1")
> ```
