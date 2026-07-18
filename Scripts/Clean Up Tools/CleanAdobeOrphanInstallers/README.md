# Adobe Orphan Installer Patch Cleaner Guide

## Overview
> **Short Description:** Cleans Adobe installer patch files.

This script scans the `C:\Windows\Installer` directory for orphaned Adobe installer patch files (`.msp`), calculates potential disk space savings, and deletes them. It runs in a dry-run (WhatIf) mode first by default to allow review before deletion.

### Key Features
* **Installer Metadata Scanning:** Queries the Windows Installer COM API to extract author and subject properties of each `.msp` patch.
* **Adobe Target Filtering:** Specifically identifies Adobe patches (e.g. Acrobat updates) and ignores unrelated system updates.
* **Service Lock Handling:** Temporarily stops the Windows Installer (`msiserver`) service to ensure files are unlocked.
* **Permission Ownership Acquisition:** Automatically takes file ownership (`takeown`) and grants permissions (`icacls`) to the administrators group to bypass file locks.

> [!NOTE]
> **Log File Location:** `C:\Logs\CleanAdobeOrphanInstallers\CleanAdobeOrphanInstallers_DDMMYY.log` (or `$env:SystemDrive\Logs\CleanAdobeOrphanInstallers\CleanAdobeOrphanInstallers_DDMMYY.log`)

## Prerequisites
OS Support: Windows 10 / 11 / Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required (elevation check included)
Dependencies: Requires the Windows Installer (`msiserver`) service.

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console with administrative privileges (Run as Administrator).
2. Execute the script to stop the installer service and scan the files. It will list all Adobe files found and calculate the total size.
3. A prompt will ask: `WhatIf dry-run completed. Do you want to run this for real now? (Y/N)`
4. Type `Y` or `Yes` and press Enter to execute the real cleanup. The script will take file ownership, grant full access, delete the files, and output space recovered.
5. Alternatively, run with the `-Force` switch to immediately clean the files without prompts.
6. The script automatically restarts the `msiserver` service upon exit.

### 2. Logging & Outputs
* Details on patch names, sizes, ownership commands, and total space recovered are displayed on the console.


## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean Up Tools/CleanAdobeOrphanInstallers/03_CleanAdobeOrphanInstallers.ps1")
> ```
