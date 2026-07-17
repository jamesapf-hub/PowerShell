# Intune Desktop Background Packager Guide

## Overview
A Windows Forms GUI utility designed to automate the building, staging, and packaging of corporate Desktop Background and Lock Screen deployments for Microsoft Intune. 

It handles local files or remote image URLs, automatically downloads the Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`), compiles the required scripts (`Install.ps1`, `ApplyBackground.ps1`, `DownloadBackgrounds.ps1`, and `Uninstall.ps1`), packages them into a `.intunewin` bundle, and generates custom detection scripts and setup configurations.

### Key Features
*   **Intuitive WinForms GUI:** Easily select and configure Desktop Background and Lock Screen management parameters.
*   **Staggered Download Mode:** Built-in sleep randomizer (staggering) for remote image URL deployments to prevent rate-limiting when thousands of corporate endpoints download branding assets simultaneously.
*   **Active Desktop Enforcements:** Prevents end-users from changing wallpaper or lock screens after deployment.
*   **Automated Packaging:** Dynamically stage required files and automatically pack the bundle using the Microsoft Content Prep tool.
*   **Auto-generated Assets:** Generates custom detection scripts (`DetectionScript.ps1`) and instructions (`InstallCommands.txt`) ready for the Intune Web Portal.

> [!NOTE]
> **Log File Location:** `C:\Logs\DesktopBackground\`

## Prerequisites
OS Support: Windows 10 / 11
PowerShell: Windows PowerShell 5.1+
Permissions: Local Administrator rights required
Dependencies: Windows Forms assemblies (built-in) and active internet connection (to download `IntuneWinAppUtil.exe` if not present).

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open an elevated PowerShell prompt (Run as Administrator).
2. Run `Create-IntuneBackgroundApp.ps1`.
3. In the GUI form:
   - Configure **Desktop Background** settings (Local file or dynamic URL).
   - Configure **Lock Screen** settings (same as desktop, a separate local file, or a separate URL).
   - Check options to lock/prevent users from modifying wallpapers.
   - Specify the target **Output Directory** for the packaged assets.
4. Click **Generate Intune Package**.
5. Once complete, check your specified output folder for:
   - `DesktopBackground_[Date].intunewin`
   - `DetectionScript.ps1`
   - `InstallCommands.txt`
6. Upload the `.intunewin` package to Intune and configure it using the parameters generated in `InstallCommands.txt` and the detection logic in `DetectionScript.ps1`.

### 2. Logging & Outputs
Staging files and execution logs are generated temporarily inside the script's root staging folder and cleaned up after compilation.
Installation and update history on endpoints are saved to:
*   `C:\Logs\DesktopBackground\Install.txt`
*   `C:\Logs\DesktopBackground\ApplyBackground.txt`
*   `C:\Logs\DesktopBackground\DownloadBackgrounds.txt` (Daily check-in task)

## Fast Execute
> [!TIP]
> **Run locally in PowerShell (as Administrator):**
> Execute the packager utility GUI locally on your workstation:
> ```powershell
> powershell.exe -ExecutionPolicy Bypass -File .\Create-IntuneBackgroundApp.ps1
> ```
