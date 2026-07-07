# Browser Cache Cleaner Guide

## Overview
This script simulates the process of cleaning browser cache files for Google Chrome and Microsoft Edge on local Windows systems.

### Key Features
* **Multi-Browser Targeting:** Targets Google Chrome and Microsoft Edge cache folders.
* **Non-Destructive Simulation:** Runs safely as a simulation, displaying status outputs to the user.

> [!NOTE]
> **Log File Location:** `C:\Logs\BrowserCacheCleaner\BrowserCacheCleaner_DDMMYY.log` (or `$env:SystemDrive\Logs\BrowserCacheCleaner\BrowserCacheCleaner_DDMMYY.log`)

## Prerequisites
OS Support: Windows 10 / 11
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Standard User context or Administrator

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open a PowerShell console.
2. Run the script.
3. The script will output startup logs, pause for 3 seconds to simulate the scan and cache delete process, and report completion.

### 2. Logging & Outputs
* Status messages for Google Chrome and Microsoft Edge cache purges are printed directly to the console.

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/BrowserCacheCleaner/02_BrowserCacheCleaner.ps1")
> ```