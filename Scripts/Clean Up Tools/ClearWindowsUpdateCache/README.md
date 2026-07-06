# Clear Windows Update Cache

---

## Overview

Stops Windows Update services (wuauserv, bits), purges the SoftwareDistribution download cache folder, and restarts the services to resolve update loop issues and free disk space.

> [!NOTE]
> **Log File Location:** `C:\Logs\ClearWindowsUpdateCache\SD_Clear_DDMMYY.log` (or `$env:SystemDrive\Logs\ClearWindowsUpdateCache\SD_Clear_DDMMYY.log`)

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
   .\04_ClearWindowsUpdateCache.ps1
   ```
3. Review the dry-run results, then type `Y` to execute the real cleanup.
