# M365 Admin Console Guide

---

## Overview

The `Console.ps1` script launches a modern, dark-themed (Kinetic Command) administration dashboard. It queries Microsoft Graph to fetch all licensed tenant users, dynamically resolves product subscription identifiers (SKU IDs), calculates inactive user periods, and provides tools to search, filter, and export data directly from your desktop.

### Key Features
* **Multi-Connection Support:** Authenticate via Delegated Graph API (M365 User Login), Application Secret Token, Direct CSV Import, or run offline using the Demo Sandbox.
* **Cost Savings Analysis:** Dynamically calculates accumulated wasted license spend and projects recurring monthly savings for inactive accounts.
* **Interactive Visualization:** Displays interactive charts showing license distributions and user inactivity timelines, along with a live telemetry execution log.
* **Advanced Exclusion Filter:** Prompts to filter/exclude Seriun or JP verification accounts during report export.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 (due to WPF graphical requirements)
* **PowerShell:** PowerShell 7.2 or later (Windows PowerShell 5.1 is not supported)
* **Permissions:** Global Reader, Global Administrator, Security Reader, or Reports Reader Entra ID roles. Graph scopes required: `User.Read.All`, `AuditLog.Read.All`, `Directory.Read.All`, `Organization.Read.All`, and `LicenseAssignment.Read.All`.
* **Dependencies:** `Microsoft.Graph` module (will auto-install if missing), `ImportExcel` module (for Excel reports; checks and prompts for permission to install on-demand).

---

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### How to Launch the GUI
* **Option A (Recommended):** Double-click the `LaunchConsole.bat` helper script. It forces PowerShell 7 to initialize in STA (Single-Threaded Apartment) mode, which is required for WPF graphics.
* **Option B (Manual):** Run the following command from PowerShell 7:
  ```powershell
  pwsh.exe -NoProfile -ExecutionPolicy Bypass -File .\Console.ps1
  ```

#### Using the Dashboard
1. **Authenticate:** Select a login method. Use *Launch Demo Sandbox* for offline testing with mock data, or *Import CSV Report* to load previous audits.
2. **Review Metrics:** View total licensed users, inactive users (30d, 90d, 1yr), and potential monthly cost savings.
3. **Filter Directory:** Navigate to the *User Directory* tab, search display names/UPNs, and filter by inactivity flags.
4. **Export Audits:** Select *Export CSV* or *Export Excel* at the bottom to save findings.

### 2. Standalone CLI Script (Headless)
For automated scheduled runs, use the headless `LicencedUsersSigninDate.ps1` script. Update it to authenticate using your client certificate details:
```powershell
Connect-MgGraph -TenantId "your-tenant-id" -ClientId "your-client-id" -CertificateThumbprint "your-thumbprint"
```

### 3. Logging & Outputs
* **GUI Telemetry:** Displayed live in the bottom telemetry log feed.
* **CLI Logging:** Appends execution logs to `C:\Seriun\log\LicensedUsers_RunLog_YYMMDD.log`.

---

## Command

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Console.ps1
```
