# Schedule Kiosk App Runtime Fix

---

## Overview

Creates a scheduled task to run Kiosk Shell App runtime fixes on user authentication to prevent logon UI hanging states.

> [!NOTE]
> **Log File Location:** `C:\Logs\Schedule-KioskAppRuntimeFix\Kiosk_Runtime_Fix.txt` (or `$env:SystemDrive\Logs\Schedule-KioskAppRuntimeFix\Kiosk_Runtime_Fix.txt`)

---

## Prerequisites

* **OS Support:** Windows 10 / 11
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.

---

## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Schedule-KioskAppRuntimeFix.ps1
   ```
