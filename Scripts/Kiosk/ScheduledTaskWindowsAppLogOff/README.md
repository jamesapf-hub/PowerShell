# Scheduled Task Windows App Log Off Watchdog

---

## Overview

Sets up a watchdog scheduled task to monitor UWP app processes. If the primary kiosk app closes, the script automatically logs off the local user to keep the terminal secure.

---

## Prerequisites

* **OS Support:** Windows 10 / 11
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.

---

## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\ScheduledTaskWindowsAppLogOff.ps1
   ```
