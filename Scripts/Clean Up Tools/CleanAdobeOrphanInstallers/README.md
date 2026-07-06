# Clean Adobe Orphan Installers

---

## Overview

Scans `C:\Windows\Installer` for orphaned Adobe Acrobat / Reader update patch files (`.msp`), calculates potential storage savings, and purges them to recover disk space.

> [!NOTE]
> **Log File Location:** `C:\Logs\CleanAdobeOrphanInstallers\CleanAdobeOrphanInstallers_DDMMYY.log` (or `$env:SystemDrive\Logs\CleanAdobeOrphanInstallers\CleanAdobeOrphanInstallers_DDMMYY.log`)

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.

---

## Walkthrough & Usage Guide

1. Open PowerShell as an Administrator.
2. Run the script:
   ```powershell
   .\03_CleanAdobeOrphanInstallers.ps1
   ```
3. Review the dry-run results showing the total space that can be recovered, then type `Y` to execute the real cleanup.
