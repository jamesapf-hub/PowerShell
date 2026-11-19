# Install Datto RMM Agent Guide

## Overview

Downloads and silently installs the **Datto RMM Agent (`CagService`)** onto a Windows endpoint using your target Datto platform subdomain and client Site UID. The script automatically verifies whether the agent is already installed, enforces TLS 1.2+, validates the downloaded binary, checks the installer exit code, confirms service registration, and cleans up temporary installation files.

### Key Features

*   **Idempotent Pre-Check:** Immediately detects if `CagService` is already installed and exits cleanly (`Exit Code 0`) to prevent duplicate installations.
*   **Parameter & RMM Variable Support:** Accepts `-Platform` and `-SiteID` command-line parameters or reads `$env:DATTO_PLATFORM` and `$env:DATTO_SITE_ID` environment variables for Group Policy, Intune, or RMM onboarding.
*   **Exit Code & Service Verification:** Launches `DRMMSetup.exe` via `Start-Process -Wait -PassThru`, validates exit code `0`, and confirms `CagService` registration post-install.
*   **Automatic Cleanup & Audit Logging:** Removes `%TEMP%\DRMMSetup.exe` in a `finally` block and writes UK-timestamped audit logs.

---

## Prerequisites

*   **OS Support:** Windows 10, Windows 11, Windows Server 2016 / 2019 / 2022 / 2025
*   **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
*   **Permissions:** Local Administrator rights ( or `NT AUTHORITY\SYSTEM` via Intune / Group Policy)
*   **Network:** Outbound HTTPS (`TCP 443`) access to `https://<Platform>.centrastage.net`

---

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Retrieve your Datto RMM **Platform name** (e.g., `pinotage`, `merlot`, `concord`, `vidal`, `zinfandel`, `syrah`) and the target client **Site ID (GUID)** from the Datto RMM portal (**Sites > All Sites > Site Settings**).
2. Run the script from an elevated PowerShell prompt passing `-Platform` and `-SiteID`:
   ```powershell
   .\Install-DattoRmmAgent.ps1 -Platform "merlot" -SiteID "12345678-abcd-1234-efab-1234567890ab"
   ```
   *(Alternatively, set `$env:DATTO_PLATFORM` and `$env:DATTO_SITE_ID` before invoking the script).*

### Logging & Diagnostics

| Log File Path | Description |
| :--- | :--- |
| `C:\Logs\Install-DattoRmmAgent\DattoAgentInstall.txt` | Timestamped (`dd/MM/yy HH:mm:ss`) log tracking pre-check status, download duration, installer exit code, and `CagService` state. |

### How Detection & Installation Work

*   **Pre-Installation Check:** Queries `Get-Service -Name "CagService"`. If present, logs the current service status and exits with code `0`.
*   **Installer Validation:** Downloads `https://<Platform>.centrastage.net/csm/profile/downloadAgent/<SiteID>` to `%TEMP%\DRMMSetup.exe` and verifies the binary size exceeds 10 KB before executing.
*   **Post-Install Verification:** Waits for `DRMMSetup.exe` to return exit code `0`, removes the temporary installer, and confirms `CagService` is registered.

---

## Fast Execute

> [!TIP]
> **Full Transparent GitHub Command:**
> ```powershell
> $env:DATTO_PLATFORM = "merlot"
> $env:DATTO_SITE_ID  = "YOUR-SITE-GUID"
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Software/Install-DattoRmmAgent/Install-DattoRmmAgent.ps1")
> ```
