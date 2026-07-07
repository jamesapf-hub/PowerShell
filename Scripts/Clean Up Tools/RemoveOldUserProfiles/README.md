# Orphaned AD Domain User Profile Cleaner Guide

## Overview
This script scans the local Windows registry for user profiles, filters out system, local, and active profiles, and identifies orphaned or deleted Active Directory domain user profiles. It deletes them through WMI to ensure clean removal from the system.

### Key Features
* **Registry Profile Scanning:** Reads profile subkeys under `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList`.
* **Domain Account Filtering:** Resolves SIDs to NTAccounts to filter out local and built-in profiles.
* **Folder Size Auditing:** Recursively measures profile folder sizes on disk.
* **Clean WMI Deletion:** Uses WMI (`Win32_UserProfile`) to safely delete profiles, clearing registry subkeys and folder paths.
* **Exclusion Support:** Excludes active accounts and supports custom exclusions via the `-ExcludeUsers` parameter.

> [!NOTE]
> **Log File Location:** `C:\Logs\RemoveOldUserProfiles\RemoveOldUserProfiles_DDMMYY.log` (or `$env:SystemDrive\Logs\RemoveOldUserProfiles\RemoveOldUserProfiles_DDMMYY.log`)

## Prerequisites
OS Support: Windows 10 / 11 / Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required (elevation check included)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console with administrative privileges (Run as Administrator).
2. Execute the script. It will scan local profiles and print out a formatted table of orphaned profiles, paths, and sizes.
3. To run with custom exclusions, pass a list of accounts: `.\05_RemoveOldUserProfiles.ps1 -ExcludeUsers "domain\user1", "domain\user2"`
4. A prompt will ask: `WhatIf dry-run completed. Do you want to run this for real now? (Y/N)`
5. Type `Y` or `Yes` and press Enter to trigger the clean WMI profile deletions.
6. Run with `-Force` to execute silently.

### 2. Logging & Outputs
* Outputs scanning statistics, user profile tables, and WMI deletion status codes.

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/RemoveOldUserProfiles/05_RemoveOldUserProfiles.ps1")
> ```