# PowerShell Scripts Repository

Welcome to your central repository for all PowerShell scripts. This repository is configured with a reusable logging structure, static code analysis (linting), and organized directories to keep your scripts clean, consistent, and well-documented.

## Repository Directory Structure

*   [Modules/Logging/Logging.psm1](Modules/Logging/Logging.psm1) - Shared PowerShell module for script logging (automatically sets up and manages `C:\Logs` directory).
*   [Templates/ScriptTemplate.ps1](Templates/ScriptTemplate.ps1) - A starter template for writing new scripts with built-in logging and error handling.
*   [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1) - Best-practice static analysis rules for script formatting and quality.
*   `Scripts/` - Folder where you can add your custom, standalone scripts.
    *   [Scripts/Networking/PSDiscovery/Get-SwitchPortInfo.ps1](Scripts/Networking/PSDiscovery/Get-SwitchPortInfo.ps1) - Automates capturing and parsing CDP/LLDP switch port packets with logging.

## Getting Started

### 1. Creating a New Script
To create a new script with automatic logging enabled, copy the starter template:
```powershell
Copy-Item "Templates/ScriptTemplate.ps1" "Scripts/YourNewScript.ps1"
```
Open the new script, modify the metadata, and write your logic inside the main `try` block.

### 2. How the Logging Works
Each script that uses the template will:
1. Automatically verify if `C:\Logs` exists. If not, it will be created.
2. Create a log file specific to that script named `[ScriptName]_[YYYY-MM-DD].log` in `C:\Logs`.
3. Use the `Write-ScriptLog` function to write timestamped messages to both the console host and the log file simultaneously.

### 3. Setting execution policy (if blocked)
If you cannot run scripts due to execution policy limits, you can set it in an administrator PowerShell session:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Script Categories & Descriptions

To keep the portal navigation meaningful, the scripts are organized into the following categories under `Scripts/`:
*   **Clean Up Tools:** Utilities to clean system caches, remove temporary files, and reclaim disk space.
*   **Configuration:** OS parameter customization, registry settings adjustments, and profile changes.
*   **Datto:** Datto RMM component scripts, policy triggers, and automated status checks.
*   **Intune:** Intune app packaging helpers, custom proactive remediations, and MDM policy templates.
*   **Kiosk:** Dedicated locking and provisioning scripts for single-app kiosk endpoints.
*   **Microsoft 365:** User creation, license checks, and tenant management scripts (Exchange, Entra, Teams).
*   **Networking:** Ping diagnostics, subnet scanners, DNS resolvers, and network status checks.
*   **Sample:** Reference templates demonstrating standard portal layouts, logging, and error handling.
*   **Software:** Silent software deployment wrappers, updates, and application uninstallation scripts.

## Static Code Analysis (Linting)

This repository includes custom settings for **PSScriptAnalyzer** to keep code formatting clean and uniform. If you use VS Code, it will automatically detect [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1) and show warnings/errors inline.

To run linting manually on all files in the repository:
```powershell
Invoke-ScriptAnalyzer -Path . -Settings .\PSScriptAnalyzerSettings.psd1 -Recurse
```
