# Entra ID SMS Deprecation Audit Guide

## Overview
> **Short Description:** Audits Entra ID user authentication methods to identify high-risk SMS dependencies before Microsoft deprecates SMS/voice MFA.

An interactive PowerShell 7 WPF tool and color-coded Excel report generator designed to audit Microsoft Entra ID user MFA registration methods. With Microsoft deprecating SMS/voice telephony as primary authentication methods, this tool helps administrators identify SMS dependencies before users are impacted.

### Key Features
* **PowerShell 7 WPF GUI:** Modern dark-theme window providing real-time scanning progress, live tenant connection status, and interactive user filtering.
* **Microsoft Graph API Audit:** Queries `/reports/authenticationMethods/userRegistrationDetails` to analyze registered methods and default sign-in preferences (`userPreferredMethod`).
* **Risk Classification:**
  * 🔴 **SMS ONLY (High Risk):** Users with only SMS/Phone registered. These users will lose access upon SMS deprecation if secondary methods are not registered.
  * 🟡 **SMS DEFAULT (Medium Risk):** Users with SMS set as default, but who have secondary methods available (such as Microsoft Authenticator or FIDO2).
  * 🟢 **SECURE MFA (Compliant):** Users defaulting to Microsoft Authenticator Push, FIDO2/Passkeys, or Software OATH TOTP.
  * ⚪ **UNREGISTERED:** Users with no MFA registered.
* **Color-Coded XLSX Export (`ImportExcel`):**
  * **Executive Summary Sheet:** High-level metrics, tenant account, timestamp, and overall SMS Dependency Ratio.
  * **SMS Deprecation Audit Sheet:** Detailed user table auto-formatted with color-coded rows (Red for SMS Only, Amber for SMS Default, Green for Secure MFA), auto-filters, and frozen headers.

## Prerequisites
OS Support: Windows 10 / 11 or Windows Server 2019 / 2022
PowerShell: PowerShell Core 7+ (`pwsh.exe`)
Permissions: Global Reader, Reports Reader, or Authentication Administrator
Entra ID Scopes Required: `User.Read.All`, `UserAuthenticationMethod.Read.All`, `Reports.Read.All`, `Directory.Read.All`
Dependencies: `Microsoft.Graph`, `ImportExcel` (automatically installed for `CurrentUser` if missing)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open PowerShell 7 (`pwsh.exe`) or double-click `launchcontrol.bat`.
2. Execute the launcher command:
   ```powershell
   pwsh.exe -ExecutionPolicy Bypass -File .\Start-MFACheckerGUI.ps1
   ```
3. Click **Connect to Entra ID** and sign in with an admin account possessing required Graph API scopes.
4. Click **Scan MFA Methods** to execute tenant analysis.
5. Review real-time metrics and filter by risk category directly within the GUI grid.
6. Click **Export Color-Coded XLSX** to generate and open the formatted Excel report workbook.

### 2. Logging & Outputs
* **GUI Live Log:** Timestamped activity logs are displayed directly in the GUI console output box.
* **Excel Workbook (`.xlsx`):** Color-coded audit report saved to the execution directory containing executive KPI summary and detailed user rows.

## Fast Execute

> [!TIP]
> **Run Locally in PowerShell 7 (as Administrator):**
> Download and extract the package folder, then execute locally:
> ```powershell
> pwsh.exe -ExecutionPolicy Bypass -File .\Start-MFACheckerGUI.ps1
> ```
