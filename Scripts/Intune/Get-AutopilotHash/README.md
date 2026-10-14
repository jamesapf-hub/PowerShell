# Get Autopilot Hash Guide

## Overview
> **Short Description:** Extracts the Windows Autopilot 4KB Hardware Hash natively via WMI/CIM and uploads it to `phnx.it` for instant short-URL CSV download on another machine (with multi-PC Batch PIN grouping).

When enrolling new or freshly reset Windows 10 and Windows 11 machines into Microsoft Intune via Windows Autopilot, engineers normally have to install `Get-WindowsAutopilotInfo` from PSGallery (which requires NuGet prompts and often fails in OOBE) and copy the resulting `.csv` file onto a physical USB flash drive.

**Get-AutopilotHash** eliminates USB drives and PSGallery dependencies completely. By running `irm phnx.it/APHASH | iex` from an elevated PowerShell prompt (or `Shift + F10` during Windows OOBE), the script extracts the BIOS serial number and `MDM_DevDetail_Ext01` hardware hash natively, saves a local backup to `C:\HWID`, uploads the record to the Cloudflare Edge, and prints a short download link (e.g., `https://phnx.it/ap/K9X4M` or `https://phnx.it/BENCH1`) that immediately downloads an Intune-ready `.csv` file when opened on your admin workstation.

### Key Features
* **Zero PSGallery / NuGet Dependencies:** Queries `root/cimv2/mdm/dmmap` (`MDM_DevDetail_Ext01`) directly via native CIM, executing in under 2 seconds even inside Windows OOBE (`Shift + F10`).
* **Instant Short-URL CSV Download:** Generates a compact `phnx.it` link (`https://phnx.it/ap/<CODE>` and `https://phnx.it/<CODE>`) that directly downloads `AutopilotHWID-<CODE>.csv` on any browser.
* **Multi-Device Batch PIN Grouping:** Enter a custom Batch PIN (e.g., `BENCH1`) across 5, 10, or 50 laptops on a build bench—every laptop using that PIN automatically appends and deduplicates into a **single combined CSV file** at `https://phnx.it/ap/BENCH1`.
* **Optional Intune Group Tag Support:** Supports tagging devices during capture so the `Group Tag` column is automatically populated for dynamic Entra ID group assignment.
* **Dual Local + Cloud Redundancy:** Always writes a local backup CSV to `C:\HWID\AutopilotHWID-<Serial>.csv` and diagnostic logs to `C:\Logs\Get-AutopilotHash`.

> [!NOTE]
> **Log File Location:** `C:\Logs\Get-AutopilotHash\AutopilotHash_[Timestamp].log` (or `%LOCALAPPDATA%\Logs\Get-AutopilotHash\AutopilotHash_[Timestamp].log`)

## Prerequisites

* **OS Support:** Windows 10 (1809+) and Windows 11 (Pro, Enterprise, Education)
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator privileges (or `SYSTEM` context via `Shift + F10` in Windows OOBE)
* **Dependencies:** None (uses built-in Windows CIM/WMI providers and HTTPS egress)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### Scenario A: Capturing a Single Machine (OOBE or Desktop)
1. On the target Windows machine:
   - **If at the Windows OOBE setup screen:** Press `Shift + F10` to open Command Prompt, then type `powershell` and press **Enter**.
   - **If on the Windows Desktop:** Right-click the Start button and select **Terminal (Admin)** or **PowerShell (Admin)**.
2. Run the short command:
   ```powershell
   irm phnx.it/APHASH | iex
   ```
3. Press **Enter** at the optional `Batch PIN` and `Group Tag` prompts (or enter a Group Tag if desired).
4. Copy or type the displayed short URL (e.g. `https://phnx.it/ap/K9X4M` or `https://phnx.it/K9X4M`) into the browser on your admin machine—the `AutopilotHWID-K9X4M.csv` file will download immediately.

#### Scenario B: Capturing Multiple Machines into One Combined CSV (Batch PIN)
1. Choose a short Batch PIN for your build session (for example, `BENCH1`).
2. On each laptop, run:
   ```powershell
   irm phnx.it/APHASH | iex
   ```
   When prompted for **Batch PIN**, type `BENCH1` (or pre-set `$env:AP_PIN = "BENCH1"; irm phnx.it/APHASH | iex` for zero prompts!).
3. Once all laptops have been scanned, open `https://phnx.it/ap/BENCH1` on your admin laptop. You will receive a single combined `AutopilotHWID-BENCH1.csv` containing every scanned laptop (automatically deduplicated by Serial Number).

### 2. Intune Portal Upload Parameters

| Setting / Field | Value / Path |
| :--- | :--- |
| **Admin Center Portal** | [Microsoft Intune Admin Center](https://intune.microsoft.com) |
| **Navigation Path** | **Devices** > **Enrollment** > **Windows** > **Windows Autopilot** > **Devices** |
| **Action** | Click **Import** in the top command bar and select your downloaded `AutopilotHWID-<CODE>.csv` file |
| **CSV Columns (Standard)** | `Device Serial Number,Windows Product ID,Hardware Hash` |
| **CSV Columns (With Group Tag)** | `Device Serial Number,Windows Product ID,Hardware Hash,Group Tag` |
| **Cloud Link Retention** | 24 Hours (auto-purged from Cloudflare D1 after expiry) |

### 3. Step-by-Step Intune Admin Center Setup

1. Open `https://phnx.it/ap/<CODE>` on your admin workstation to download `AutopilotHWID-<CODE>.csv`.
2. Sign in to the **Microsoft Intune Admin Center** (`https://intune.microsoft.com`).
3. Navigate to **Devices** > **Windows** > **Enrollment** > **Devices** (under *Windows Autopilot*).
4. Click **Import** at the top of the blade.
5. Browse to and select your downloaded `AutopilotHWID-<CODE>.csv` file.
6. Confirm that Intune reports formatting validation passed, and click **Import** (sync typically completes within 5–15 minutes).

### Logging & Diagnostics

| Output File | Path | Description |
| :--- | :--- | :--- |
| **Execution Log** | `C:\Logs\Get-AutopilotHash\AutopilotHash_<ddMMyy-HHmmss>.log` | Timestamped audit log of CIM queries, serial extraction, and cloud upload response. |
| **Local CSV Backup** | `C:\HWID\AutopilotHWID-<SerialNumber>.csv` | Local copy of the Intune Autopilot CSV saved on the target machine before cloud upload. |

### How Detection & Extraction Work

Instead of downloading third-party scripts from the PowerShell Gallery, `Get-AutopilotHash.ps1` queries the local Windows Management Instrumentation (`root/cimv2/mdm/dmmap`) namespace directly:
- **Serial Number:** Extracted from `Win32_BIOS.SerialNumber`.
- **Hardware Hash:** Extracted from `MDM_DevDetail_Ext01.DeviceHardwareData` filtered on `InstanceID='Ext' AND ParentID='./DevDetail'`.
- **Batch Deduplication:** If you run the script twice on the same machine using the same Batch PIN, the cloud endpoint updates the existing serial number entry rather than duplicating rows in the CSV.

## Fast Execute

> [!TIP]
> **Short Branded URL (phnx.it):**
> ```powershell
> irm phnx.it/APHASH | iex
> ```
>
> **Zero-Prompt Batch Mode (Pre-setting Batch PIN & Group Tag):**
> ```powershell
> $env:AP_PIN="BENCH1"; $env:AP_TAG="StandardUser"; irm phnx.it/APHASH | iex
> ```
>
> **Full Transparent GitHub Command:**
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Intune/Get-AutopilotHash/Get-AutopilotHash.ps1")
> ```
