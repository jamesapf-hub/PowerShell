# Switch Information Extractor (`Get-SwitchInfo`)


> [!NOTE]
> **Switch Audit Log Location:** `C:\Logs\PSDiscovery\get-switchinfo.log` (or `$env:SystemDrive\Logs\PSDiscovery\get-switchinfo.log`)

This utility retrieves VLAN assignments, audits port configurations, and checks active switch status profiles.

## Prerequisites
**OS Support:** Windows / macOS / Linux
**PowerShell:** PowerShell Core `v7.0` or higher
**Permissions:** Standard User context (no local Windows administrator rights are required). Switch read access credentials/community strings are required.
**Network Ports:** Local network access to switch port interfaces (e.g. SNMP Port 161 or SSH Port 22) must be open.

## Walkthrough & Usage
1. Connect switch management console IP inside target script parameters.
2. Click **Copy Run Command** on the right side panel.
3. Open an Elevated PowerShell terminal shell window.
4. Execute the command to launch configuration extraction telemetry.
5. Review the resulting status tables and audit parameters output file written to:
   `$env:SystemDrive\Logs\PSDiscovery\get-switchinfo.log`


## Folder Contents
- **`get-switchinfo.ps1`**: Main executable script file.
- **`README.md`**: Walkthrough information documentation page.

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Networking/PSDiscovery/get-switchinfo.ps1")
> ```