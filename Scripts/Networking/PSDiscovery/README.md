# Switch Information Extractor (`Get-SwitchInfo`)

This utility retrieves VLAN assignments, audits port configurations, and checks active switch status profiles.

## Prerequisites
* PowerShell Core `v7.0` or higher
* Administrator execution permission parameters
* Open local port interfaces configured

## Walkthrough & Usage
1. Connect switch management console IP inside target script parameters.
2. Click **Copy Run Command** on the right side panel.
3. Open an Elevated PowerShell terminal shell window.
4. Execute the command to launch configuration extraction telemetry.
5. Review the resulting status tables and audit parameters output file written to:
   `C:\Seriun\log\get-switchinfo.log`

---

## Folder Contents
- **`get-switchinfo.ps1`**: Main executable script file.
- **`README.md`**: Walkthrough information documentation page.
