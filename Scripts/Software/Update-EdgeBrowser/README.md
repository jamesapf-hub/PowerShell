# Update Microsoft Edge Browser


## Overview

Downloads the latest Microsoft Edge Enterprise x64 MSI package directly from Microsoft and updates/installs the browser silently.


> [!NOTE]
> **Log File Location:** `C:\Logs\Update-EdgeBrowser\EdgeUpdate.txt` (or `$env:SystemDrive\Logs\Update-EdgeBrowser\EdgeUpdate.txt`)

## Prerequisites

**OS Support:** Windows 10 / 11 / Windows Server
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
**Permissions:** Local Administrator rights required.


## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Update-EdgeBrowser.ps1
   ```

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Software/Update-EdgeBrowser/Update-EdgeBrowser.ps1")
> ```
