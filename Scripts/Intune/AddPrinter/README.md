# AddPrinter

## Overview
This is an interactive desktop GUI packaging tool to prepare, structure, and compile TCP/IP printer installation packages into Microsoft Intune Win32 Apps (.intunewin). It automatically generates the required detection rules, install/uninstall scripts, and configuration files without requiring manual coding.

## Prerequisites
- Windows PowerShell 5.1 (Pre-installed on all Windows systems) or PowerShell Core 7+
- Administrator Execution Policy Bypass
- System PresentationFramework assemblies (for the WPF user interface)
- Official Microsoft `IntuneWinAppUtil.exe` in the same directory (the tool will attempt to download it automatically if missing)

## Walkthrough
1. Launch the application by executing `AddPrinter.ps1` (or using the `Start-Gui.bat` launcher).
2. Enter the target **Printer Name**, exact **Driver Name**, and static **Printer IP Address**.
3. Select the **Printer Driver Folder** (which must contain the driver `.inf` file) and choose an **Output Folder**.
4. Click **Build IntuneWin Package** to run pre-verifications, compile the `.intunewin` package, and generate custom detection and instruction scripts.

## Command
powershell.exe -ExecutionPolicy Bypass -File .\AddPrinter.ps1
