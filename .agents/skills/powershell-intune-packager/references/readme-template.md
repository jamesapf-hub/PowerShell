# PowerShell Tool `README.md` & STA Reference (`references/readme-template.md`)

## 1. WPF Single-Threaded Apartment (`STA`) Check
```powershell
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
    exit
}
```

## 2. Required `README.md` Section Order
1. `# <Tool Name> Guide` + immediate 1–2 sentence overview paragraph (parsed by `sync_baselines.cjs` into `metadata.json`).
2. `## Overview` + `### Key Features`
3. `## Prerequisites`
4. `## Walkthrough & Usage Guide` (`### 1. Step-by-Step Instructions`, optional `### 2. Intune Portal Upload Parameters` table & `### 3. Step-by-Step Intune Admin Center Setup`, `### Logging & Diagnostics`, `### How Detection & Installation Work`)
5. `## Fast Execute`:
   - **Single-File Remote Script**:
     ````markdown
     > [!TIP]
     > **Short Branded URL (phnx.it):**
     > ```powershell
     > irm https://phnx.it/<CODE> | iex
     > ```
     > **Full Transparent GitHub Command:**
     > ```powershell
     > iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/<Category>/<ToolName>/<ToolName>.ps1")
     > ```
     ````
   - **Multi-File / GUI Package**:
     ````markdown
     > [!WARNING]
     > **Bundled Package Notice:** Multi-file package; clone/download locally rather than `irm | iex`.
     >
     > [!TIP]
     > Double-click `Start-Gui.bat` or run `powershell.exe -ExecutionPolicy Bypass -File .\<ToolName>.ps1`
     ````

## 3. Canonical Intune Win32 Deployment `.txt` Output Template
Whenever an Intune Win32 packager generates a `.intunewin` file and custom detection `.ps1` script, it MUST also generate a deployment instructions `.txt` file matching the exact step-by-step order of the Microsoft Intune "Add Windows app (Win32)" wizard, with the copy-pasteable `Description` box listing `C:\Logs\<PackageFolder>\` and all log files inside that folder:

```text
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : <dd/MM/yy HH:mm:ss>
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
<PackageFileName>.intunewin


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
<App Display Name>

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
<1-2 sentence summary of what the package deploys/configures>
All runtime logs are saved to C:\Logs\<PackageFolder>\

PACKAGE DETAILS:
- <Key Detail 1> : <Value>
- <Key Detail 2> : <Value>
- Author         : JP
- App Version    : <Version>
- Created        : <dd/MM/yy>

WORKSTATION LOG LOCATION:
C:\Logs\<PackageFolder>\

LOG FILES IN THIS FOLDER:
- Install Log   : C:\Logs\<PackageFolder>\<InstallLog>.log
- Detection Log : C:\Logs\<PackageFolder>\<DetectionLog>.log
- Uninstall Log : C:\Logs\<PackageFolder>\<UninstallLog>.log
======================================================================

Publisher:
<PublisherName>

App Version:
<Version>


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\<InstallScript>.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\<UninstallScript>.ps1

Install behavior:
System

Device restart behavior:
No specific action


STEP 4: REQUIREMENTS
------------------------------------------------------------------------
Operating system architecture:
64-bit

Minimum operating system:
Windows 10 1909


STEP 5: DETECTION RULES
------------------------------------------------------------------------
Rules format:
Use a custom detection script

Script file:
<DetectionScript>.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No
========================================================================
```
