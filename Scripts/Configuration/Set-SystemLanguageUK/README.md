# Set System Language and Keyboard Locales (UK English)

Configures regional settings, user locale, system locale, and keyboard layouts to United Kingdom English (en-GB).

> [!NOTE]
> **Log File Location:** `C:\Logs\Set-SystemLanguageUK\LanguageSetup.txt` (with automatic fallback to `%LOCALAPPDATA%\Logs\Set-SystemLanguageUK\LanguageSetup.txt` if `C:\Logs` is not writable)

## Overview

`Set-SystemLanguageUK.ps1` standardizes Windows endpoints onto `en-GB` across the current user, non-Unicode system locale, Welcome/logon screen, and new-user profile templates.

### Key Features

* **Fast Core Install by Default (`-ExcludeFeatures`):** Pre-checks `Get-InstalledLanguage` to skip unnecessary downloads when `en-GB` is already present, and installs only the core `en-GB` language pack by default to avoid long waits and optional Feature on Demand (FoD) timeouts.
* **Optional Features on Demand (`-IncludeFeatures`):** Supports `-IncludeFeatures` (or `DATTO_INCLUDE_FEATURES=true` in RMM runs) to install Speech, Handwriting, Text-to-Speech, and OCR, with automatic fallback to `-ExcludeFeatures` if an optional FoD fails.
* **Full UK System Alignment:** Sets System Preferred UI Language (`Set-SystemPreferredUILanguage`), User Culture (`en-GB`), System Locale (`en-GB`), Home Location (`GeoID 242`), UK keyboard/language list, and UI override, disables cloud language sync, and copies settings to the Welcome Screen and New User templates.
* **Safe Dry-Run (`-WhatIf`) & Diagnostics:** Supports `-WhatIf` without creating logs or modifying settings, and decodes CBS/Windows Update HRESULT codes (such as `0x800F0993`) while checking WSUS, pending reboot, and Single Language edition restrictions.

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required for system-wide and language pack changes.

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Open an **Elevated PowerShell** console (Run as Administrator).
2. Run the script in fast core mode (recommended):
   ```powershell
   .\Set-SystemLanguageUK.ps1
   ```
3. To also request optional Features on Demand (Speech, Handwriting, TTS, OCR):
   ```powershell
   .\Set-SystemLanguageUK.ps1 -IncludeFeatures
   ```
4. Preview intended changes without modifying the endpoint:
   ```powershell
   .\Set-SystemLanguageUK.ps1 -WhatIf
   ```

### 2. Troubleshooting `ErrorCode: -2146498157` (`0x800F0993`)

If an endpoint previously reported `Language pack or features could only be partially installed. ErrorCode: -2146498157` (`0x800F0993`, Component-Based Servicing error):

1. **Check whether the core `en-GB` display pack already installed:**
   ```powershell
   Get-InstalledLanguage | Where-Object { $_.LanguageId -eq 'en-GB' }
   ```
   If `LanguagePacks` shows `LpCab` or `LXP`, the core UK display language is already present and only an optional FoD failed.
2. **If `LanguagePacks` is `None` or `en-GB` is missing:**
   * **Check WSUS policy:** If `HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer` is `1`, enable **Specify settings for optional component installation and component repair** (`RepairContentServerSource = 2`) so Windows can fetch language packs directly from Microsoft Update.
   * **Reboot & repair component store:** Reboot if updates are pending, run `DISM /Online /Cleanup-Image /RestoreHealth`, reboot again, then re-run `Set-SystemLanguageUK.ps1`.

## Fast Execute

> [!TIP]
> **Short Branded URL (phnx.it):**
> ```powershell
> irm https://phnx.it/UKLANG1 | iex
> ```
>
> **Full Transparent GitHub Command:**
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Configuration/Set-SystemLanguageUK/Set-SystemLanguageUK.ps1")
> ```
