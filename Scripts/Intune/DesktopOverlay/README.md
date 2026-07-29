# Intune IT Support Desktop Overlay Guide

## Overview
> **Short Description:** Visual GUI studio and automated packager to build, preview, and deploy customized desktop IT support overlays via Intune.

A comprehensive Windows Forms GUI studio and automated Intune Win32 packaging tool that enables IT administrators to design, customize, preview, and package a persistent desktop IT support overlay for corporate workstations. 

The overlay displays critical helpdesk contact information (Phone, Email), system diagnostics (Computer Name, User, IP Address, Serial Number), and brand customization directly on the user's desktop, with options for position, transparency, and top-level pin controls.

### Key Features
* **Visual Studio GUI (`OverlayStudio.ps1`):** Real-time interactive preview for customizing overlay placement, background opacity, accent colors, helpdesk phone, email, and system diagnostic data fields.
* **Instant Sandbox & Live Preview:** Launches immediate interactive preview overlays without requiring Intune enrollment or system reboots.
* **Automated Intune Win32 Packager:** Downloads the Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`) and compiles `Install-ITOverlay.ps1`, `Detect-ITOverlay.ps1`, `Start-ITOverlay.ps1`, and `Uninstall-ITOverlay.ps1` into a `.intunewin` package.
* **Persistent Logon Trigger:** Registers a lightweight Windows Scheduled Task under `C:\ProgramData\ITSupportOverlay` to automatically display the overlay at user logon.

> [!NOTE]
> **Log File Location:** `C:\ProgramData\ITSupportOverlay\install.log` and `detection.log`

## Prerequisites

**OS Support:** Windows 10 / Windows 11  
**PowerShell:** Windows PowerShell 5.1+  
**Permissions:** Local Administrator rights required to build and deploy Intune packages  
**Dependencies:** Built-in Windows Forms & System.Drawing assemblies; `IntuneWinAppUtil.exe` (included or automatically downloaded)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### Option A: Using the Visual Studio GUI (`OverlayStudio.ps1`)
1. Right-click **`Launch-Studio.bat`** and select **Run as Administrator** (or run `powershell.exe -ExecutionPolicy Bypass -File .\OverlayStudio.ps1`).
2. Customize the overlay settings in the left panel:
   - **Helpdesk Information:** Edit Helpdesk Title, Phone Number, and Support Email.
   - **Position & Styling:** Choose screen position (Bottom Right, Bottom Left, Top Right, Top Left), Opacity (transparency), Background, and Accent colors.
   - **Display Attributes:** Toggle visible fields (Computer Name, Logged-in User, IP Address, Serial Number).
3. Click **Preview Live Overlay** to view the live floating overlay widget on your screen.
4. Click **Build Intune Package (.intunewin)** to compile the deployment package.

#### Option B: Building via Command Line (`Build-IntunePackage.ps1`)
1. Open an elevated PowerShell session in the `DesktopOverlay` directory.
2. Run `.\Build-IntunePackage.ps1`.
3. The script verifies `src/` payload files, downloads `IntuneWinAppUtil.exe` if missing, and outputs the `.intunewin` package into the `output/` directory.

### 2. Intune Portal Upload Parameters

When creating the Win32 App in Microsoft Intune Admin Center:
- **Install Command:** `powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File .\Install-ITOverlay.ps1`
- **Uninstall Command:** `powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File .\Uninstall-ITOverlay.ps1`
- **Install Behavior:** System
- **Detection Rules:** Use custom script `src/Detect-ITOverlay.ps1`

## Fast Execute

> [!TIP]
> **Run Studio Locally (as Administrator):**
> Execute the Overlay Studio GUI on your workstation:
> ```powershell
> powershell.exe -ExecutionPolicy Bypass -File .\OverlayStudio.ps1
> ```
