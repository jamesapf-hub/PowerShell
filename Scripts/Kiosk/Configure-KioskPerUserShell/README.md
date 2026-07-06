# Kiosk Per-User Shell Customization

---

## Overview

The `Configure-KioskPerUserShell.ps1` script configures registry settings to launch a custom user shell (e.g. Wyze Easy Setup Shell) for a specific local kiosk user.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required (modifies Winlogon registry keys).

---

## Walkthrough & Usage Guide

1. Open an Elevated PowerShell console.
2. Execute the script specifying the kiosk username and path to your shell executable:
   ```powershell
   .\Configure-KioskPerUserShell.ps1 -KioskUser "User" -ShellPath "C:\Program Files\Wyse\WyseEasySetup\WyseEasySetupShell.exe"
   ```
