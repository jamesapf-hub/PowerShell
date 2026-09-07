# Microsoft Teams Custom Background Win32 App Packager Guide

## Overview
A PowerShell WPF GUI application designed to package custom Microsoft Teams video meeting background images into an enterprise-grade Microsoft Intune Win32 application (`.intunewin`).

The utility automates the entire lifecycle: image selection with thumbnail previews, resolution and aspect ratio validation (flagging 1920x1080 / 16:9), dynamic installer/uninstaller and custom detection script generation, silent compilation via the Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`), and provides all exact Intune configuration settings ready for deployment.

### Key Features
*   **Modern WPF GUI:** Clean Windows 11 Fluent dark theme interface.
*   **Image Management:** Select multiple PNG, JPG, or JPEG images with live thumbnail previews, file size calculations, and resolution validation (optimally 1920x1080 / 16:9).
*   **Comprehensive Teams Support:**
    *   **New Microsoft Teams (MSTeams / v2):** Deploys directly to `%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads`
    *   **Classic Microsoft Teams (v1):** Deploys to `%APPDATA%\Microsoft\Teams\Backgrounds\Uploads`
*   **Flexible & Smart Deployment Scopes:**
    *   **Smart Auto-Detect (Recommended):** Automatically evaluates the runtime identity. If run under `NT AUTHORITY\SYSTEM` (Intune System Context, Windows Sandbox test runners, SCCM), it automatically enumerates and provisions all user profiles under `C:\Users` (including `WDAGUtilityAccount` and `C:\Users\Default`), preventing the `systemprofile` trap. If executed by a logged-in user, it deploys directly into their active profile.
    *   **User Context:** Configured for Intune per-user assignment (with fail-safe multi-profile fallback if run under SYSTEM).
    *   **System Context (All Users):** Explicitly enumerates all user profiles under `C:\Users` on every run.
*   **Enterprise Diagnostics & Logging:**
    *   Automatically writes detailed timestamped logs to `C:\Log\<AppName>\Install.log`, `C:\Log\<AppName>\Detection.log`, and `C:\Log\<AppName>\Uninstall.log`.
    *   Configures NTFS permissions on `C:\Log` so standard users can write diagnostic records.
*   **Automated Win32 Packaging:** Automatically detects local copies or downloads the official `IntuneWinAppUtil.exe` and compiles the `.intunewin` package.
*   **Intune Guidance & Detection Rules:**
    *   Provides exact install and uninstall commands with one-click copy buttons.
    *   Generates a standalone custom PowerShell detection script (`Detect-TeamsBackgrounds.ps1`) alongside the package.
    *   Saves an `Intune_Configuration_Guide.txt` summary in the output directory.

## Prerequisites
*   **OS Support:** Windows 10 (1909+) or Windows 11
*   **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+ (automatically handles STA mode)
*   **Permissions:** Local Administrator rights required to build and deploy packages
*   **Dependencies:** Windows Forms and WPF Presentation assemblies (built into Windows), `tools/IntuneWinAppUtil.exe` (bundled or auto-downloaded)

---

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open an elevated PowerShell prompt or double-click `Start-Gui.bat`.
2. In the GUI form:
   - Configure **Application Name** (e.g. `Teams Backgrounds - Corporate Branding`).
   - Configure **Publisher** (e.g. `Corporate IT`) and **Version** (e.g. `1.0.0`).
   - Keep both **New Microsoft Teams** and **Classic Microsoft Teams** selected.
   - Choose **Deployment Context**: **Smart Auto-Detect (Recommended)**.
   - Verify the **Log Directory** (`C:\Log\<AppName>`) and choose an **Output Directory**.
3. Click **➕ Add Images...** to select corporate background files (`.png`, `.jpg`, `.jpeg`).
   - Review thumbnail preview, image resolution, and aspect ratio.
4. Click **🚀 Build .intunewin Package**.
5. The packager stages the files, generates `Install-TeamsBackgrounds.ps1`, `Uninstall-TeamsBackgrounds.ps1`, and `Detect-TeamsBackgrounds.ps1`, compiles the `.intunewin` bundle, and switches to the **📋 Intune Deployment Details** tab.
6. Check your output folder for:
   - `<YourAppName>.intunewin`
   - `Detect-TeamsBackgrounds.ps1`
   - `Intune_Configuration_Guide.txt`

---

### 2. Intune Portal Upload Parameters

When creating the Win32 App in Microsoft Intune Admin Center:

| Setting | Value |
| :--- | :--- |
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Install-TeamsBackgrounds.ps1"` |
| **Uninstall command** | `powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Uninstall-TeamsBackgrounds.ps1"` |
| **Install behavior** | **User** *(or **System** if deploying system-wide)* |
| **Device restart behavior** | **Determine behavior based on return codes** |
| **Return codes** | Code `0` = Success |
| **Detection rules** | **Use a custom detection script** -> Upload `Detect-TeamsBackgrounds.ps1` |
| **Run script as 32-bit process** | **No** |
| **Enforce script signature check** | **No** |

---

### 3. Step-by-Step Intune Admin Center Setup
1. Sign in to the [Microsoft Intune admin center](https://intune.microsoft.com/).
2. Navigate to **Apps** > **Windows** > **Add**.
3. Select **App type**: **Windows app (Win32)** and click **Select**.
4. Click **Select app package file** and upload the generated `.intunewin` file from your output folder.
5. In **App information**, specify Name, Description, and Publisher.
6. In **Program**, enter the Install and Uninstall commands from the table above.
7. In **Requirements**, select both **64-bit** and **32-bit** architecture and minimum OS **Windows 10 1909**.
8. In **Detection rules**, select **Use a custom detection script** and upload `Detect-TeamsBackgrounds.ps1`.
9. In **Assignments**, target your required Entra ID user or device groups.

---

### 4. Logging & Diagnostics (`C:\Log\<AppName>`)

All generated scripts and the packager tool include enterprise-grade logging outputting to `C:\Log\<AppName>\`:

| Component | Log File Path | Description |
| :--- | :--- | :--- |
| **Installation** | `C:\Log\<AppName>\Install.log` | Records host, identity, context, each file copied with size, directory creation, marker file generation, and exit status. |
| **Detection** | `C:\Log\<AppName>\Detection.log` | Records each background image checked, upload path validation, manifest verification, and exit code (`0` = Installed, `1` = Not Installed). |
| **Uninstallation** | `C:\Log\<AppName>\Uninstall.log` | Records each file deleted, manifest cleanup, and uninstallation status. |
| **Packager Tool** | `C:\Log\TeamsBackgroundPackager\Packager.log` | Records package builds, image counts, IntuneWinAppUtil compilation output, and output package path. |

> [!NOTE]
> When running under SYSTEM context, the installer automatically grants `BUILTIN\Users` modify permissions on `C:\Log` so future user-context scripts can also write diagnostic logs. If a standard user runs in an environment where `C:\Log` cannot be created, it automatically falls back gracefully to `%LOCALAPPDATA%\Log\<AppName>\`.

---

### 5. How Detection & Installation Work

*   **Installation:** Copies background images to the respective Teams `Uploads` folders and writes a persistent deployment manifest to `%LOCALAPPDATA%\TeamsCustomBackgrounds\installed.json` (user) or `%ProgramData%\TeamsCustomBackgrounds\installed.json` (system).
*   **Detection:** Intune executes `Detect-TeamsBackgrounds.ps1`. It verifies all required images are present in the target folders and validates the manifest. If all files exist, it outputs success and exits with code `0`. If any file is missing, it exits with code `1`.
*   **Uninstallation:** Removes **only** the corporate background images deployed by this specific package, leaving personal user backgrounds untouched.

---

## Fast Execute

> [!WARNING]
> **Bundled Package Notice:**
> This tool is a multi-file WPF application package (`TeamsBackgroundPackager.ps1`, `Start-Gui.bat`, `tools/IntuneWinAppUtil.exe`) and cannot be executed in-memory via `irm | iex`. Please download or clone the repository locally before executing.

> [!TIP]
> **Run locally in PowerShell (as Administrator):**
> Double-click `Start-Gui.bat` or run from an elevated console:
> ```powershell
> powershell.exe -ExecutionPolicy Bypass -File .\TeamsBackgroundPackager.ps1
> ```
