# Toggle Privacy Mode Guide

## Overview
Toggles the 'PrivacyMode' setting in the CentraStage (Datto RMM) CagService user.config file and schedules a service restart.

### Key Features
* **Dynamic Configuration Location:** Dynamically determines the correct path to the `user.config` file by querying the active DisplayVersion from the registry.
* **Non-Blocking Service Restart:** Employs a self-deleting scheduled task to restart the Datto RMM service after one minute, preventing the script execution from hanging or stalling deployment tools.
* **Robust Logging:** Records all checks, modifications, and task scheduling statuses to a local log file.

> [!NOTE]
> **Log File Location:** `$env:SystemDrive\Logs\TogglePrivacyMode\TogglePrivacyMode.log`

## Prerequisites
OS Support: Windows 10 / 11 or Windows Server
PowerShell: Windows PowerShell 5.1+
Permissions: Local Administrator rights required
Execution Policy: Bypass or RemoteSigned
Dependencies: Datto RMM agent (CentraStage) installed

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open an elevated PowerShell prompt (Run as Administrator).
2. Execute the script `TogglePrivacyMode.ps1`.
3. The script will automatically locate the Datto RMM configuration file and toggle the `PrivacyMode` value.
4. It will then register a temporary Scheduled Task named "Restart Cags Service".
5. Wait approximately 1 minute for the scheduled task to execute and restart the Datto RMM service in the background.

### 2. Logging & Outputs
The script provides color-coded console output and writes to a local log file at `$env:SystemDrive\Logs\TogglePrivacyMode\TogglePrivacyMode.log`.
- **Cyan (INFO):** Standard process steps.
- **Green (SUCCESS):** Configuration applied or task scheduled successfully.
- **Red (ERROR):** Any failure to locate the config, modify the XML, or create the scheduled task.


## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Datto/TogglePrivacyMode/TogglePrivacyMode.ps1")
> ```
