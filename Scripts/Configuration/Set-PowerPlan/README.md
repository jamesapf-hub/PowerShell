# Set Power Plan Configuration


## Overview

Configures a high-performance power plan scheme. It disables sleep mode, standby, monitor sleep timers, and USB selective suspend to prevent terminal locks.


> [!NOTE]
> **Log File Location:** `C:\Logs\Set-PowerPlan\PowerPlanConfig.txt` (or `$env:SystemDrive\Logs\Set-PowerPlan\PowerPlanConfig.txt`)

## Prerequisites

**OS Support:** Windows 10 / 11 / Windows Server
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Set-PowerPlan.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Configuration/Set-PowerPlan/Set-PowerPlan.ps1")
> ```
