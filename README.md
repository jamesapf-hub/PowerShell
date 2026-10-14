# PowerShell Scripts Repository

Welcome to your central repository for all PowerShell scripts. This repository is configured with a reusable logging structure, static code analysis (linting), and organized directories to keep your scripts clean, consistent, and well-documented.

## Repository Directory Structure

*   [Modules/Logging/Logging.psm1](Modules/Logging/Logging.psm1) - Shared PowerShell module for script logging (automatically sets up and manages `C:\Logs` directory).
*   [Templates/ScriptTemplate.ps1](Templates/ScriptTemplate.ps1) - A starter template for writing new scripts with built-in logging and error handling.
*   [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1) - Best-practice static analysis rules for script formatting and quality.
*   `Scripts/` - Categorized standalone PowerShell scripts, Win32 Intune packagers, and GUI utilities:
    *   **Azure**
        *   [Scripts/Azure/HelpfulTools/Get-AzureVmDetails/Get-AzureVmDetails.ps1](Scripts/Azure/HelpfulTools/Get-AzureVmDetails/Get-AzureVmDetails.ps1) - Queries the Azure Instance Metadata Service (IMDS) and exports VM details.
    *   **Clean Up Tools**
        *   [Scripts/Clean Up Tools/CleanAdobeOrphanInstallers/03_CleanAdobeOrphanInstallers.ps1](Scripts/Clean%20Up%20Tools/CleanAdobeOrphanInstallers/03_CleanAdobeOrphanInstallers.ps1) - Removes orphaned Adobe Acrobat/Reader `.msi` and `.msp` patch files from `C:\Windows\Installer`.
        *   [Scripts/Clean Up Tools/ClearWindowsUpdateCache/04_ClearWindowsUpdateCache.ps1](Scripts/Clean%20Up%20Tools/ClearWindowsUpdateCache/04_ClearWindowsUpdateCache.ps1) - Stops update services and purges `SoftwareDistribution\Download` cache to reclaim disk space.
        *   [Scripts/Clean Up Tools/RemoveOldUserProfiles/05_RemoveOldUserProfiles.ps1](Scripts/Clean%20Up%20Tools/RemoveOldUserProfiles/05_RemoveOldUserProfiles.ps1) - Removes stale local user profiles inactive beyond a configurable threshold.
        *   [Scripts/Clean Up Tools/SystemTempCleaner/01_SystemTempCleaner.ps1](Scripts/Clean%20Up%20Tools/SystemTempCleaner/01_SystemTempCleaner.ps1) - Cleans Windows Temp, user Temp, Prefetch, and Recycle Bin directories.
        *   [Scripts/Clean Up Tools/TidyUpWindows11/06_TidyUpWindows11.ps1](Scripts/Clean%20Up%20Tools/TidyUpWindows11/06_TidyUpWindows11.ps1) - Debloats Windows 11 consumer apps and optimizes taskbar and explorer defaults.
        *   [Scripts/Clean Up Tools/UltimateCleanUp/Start-UltimateCleanUp.ps1](Scripts/Clean%20Up%20Tools/UltimateCleanUp/Start-UltimateCleanUp.ps1) - All-in-one disk reclamation utility combining temp cleanup, update cache purge, and DISM component cleanup.
    *   **Configuration**
        *   [Scripts/Configuration/Set-PowerPlan/Set-PowerPlan.ps1](Scripts/Configuration/Set-PowerPlan/Set-PowerPlan.ps1) - Configures active Windows power scheme, sleep timeouts, and hibernation behavior.
        *   [Scripts/Configuration/Set-SystemLanguageUK/Set-SystemLanguageUK.ps1](Scripts/Configuration/Set-SystemLanguageUK/Set-SystemLanguageUK.ps1) - Enforces `en-GB` system locale, timezone, keyboard layout, and regional settings across all profiles.
        *   [Scripts/Configuration/Set-VncPassword/Set-VncPassword.ps1](Scripts/Configuration/Set-VncPassword/Set-VncPassword.ps1) - Configures encrypted TightVNC / UltraVNC registry passwords silently on managed endpoints.
    *   **Datto**
        *   [Scripts/Datto/TogglePrivacyMode/TogglePrivacyMode.ps1](Scripts/Datto/TogglePrivacyMode/TogglePrivacyMode.ps1) - Toggles Datto RMM Agent remote session privacy mode on or off.
    *   **Intune**
        *   [Scripts/Intune/AddPrinter/AddPrinter.ps1](Scripts/Intune/AddPrinter/AddPrinter.ps1) - GUI & Win32 Intune packager for deploying TCP/IP network printers and drivers.
        *   [Scripts/Intune/DesktopBackground/Create-IntuneBackgroundApp.ps1](Scripts/Intune/DesktopBackground/Create-IntuneBackgroundApp.ps1) - Packages corporate desktop and lockscreen wallpapers into an `.intunewin` deployment package.
        *   [Scripts/Intune/DesktopOverlay/OverlayStudio.ps1](Scripts/Intune/DesktopOverlay/OverlayStudio.ps1) - WPF Studio & Intune Win32 packager for deploying live hostname/IP/serial desktop watermark overlays.
        *   [Scripts/Intune/Get-AutopilotHash/Get-AutopilotHash.ps1](Scripts/Intune/Get-AutopilotHash/Get-AutopilotHash.ps1) - Windows Autopilot Hardware Hash (HWID) Harvester & `phnx.it` Cloud CSV Drop with Batch PIN support.
        *   [Scripts/Intune/TeamsBackground/TeamsBackgroundPackager.ps1](Scripts/Intune/TeamsBackground/TeamsBackgroundPackager.ps1) - Microsoft Teams Custom Background Win32 App Packager for New Teams (`MSTeams`).
    *   **Kiosk**
        *   [Scripts/Kiosk/Configure-KioskPerUserShell/Configure-KioskPerUserShell.ps1](Scripts/Kiosk/Configure-KioskPerUserShell/Configure-KioskPerUserShell.ps1) - Configures per-user custom shell launcher policies for dedicated kiosk accounts.
        *   [Scripts/Kiosk/Deploy-KioskPoliciesAndCacheReset/Deploy-KioskPoliciesAndCacheReset.ps1](Scripts/Kiosk/Deploy-KioskPoliciesAndCacheReset/Deploy-KioskPoliciesAndCacheReset.ps1) - Deploys kiosk lockdown policies and automated session cache reset tasks.
        *   [Scripts/Kiosk/Install-WindowsAppLTSC/Install-WindowsAppLTSC.ps1](Scripts/Kiosk/Install-WindowsAppLTSC/Install-WindowsAppLTSC.ps1) - Installs the Microsoft Windows App (AVD / Windows 365 client) and dependencies on Windows LTSC kiosks.
        *   [Scripts/Kiosk/Remove-UnifiedWriteFilter/Remove-UnifiedWriteFilter.ps1](Scripts/Kiosk/Remove-UnifiedWriteFilter/Remove-UnifiedWriteFilter.ps1) - Disables servicing locks and removes Unified Write Filter (UWF) protection cleanly.
        *   [Scripts/Kiosk/Schedule-KioskAppRuntimeFix/Schedule-KioskAppRuntimeFix.ps1](Scripts/Kiosk/Schedule-KioskAppRuntimeFix/Schedule-KioskAppRuntimeFix.ps1) - Registers a proactive scheduled task to repair kiosk UWP/AppX runtime registration issues.
        *   [Scripts/Kiosk/ScheduledTaskWindowsAppLogOff/ScheduledTaskWindowsAppLogOff.ps1](Scripts/Kiosk/ScheduledTaskWindowsAppLogOff/ScheduledTaskWindowsAppLogOff.ps1) - Configures automatic session logoff when the Windows App remote desktop client closes.
    *   **Microsoft 365**
        *   [Scripts/Microsoft 365/DeviceQuarantineRelease/DeviceQuarantineReleaseApp.ps1](Scripts/Microsoft%20365/DeviceQuarantineRelease/DeviceQuarantineReleaseApp.ps1) - WPF GUI utility to audit and release quarantined Exchange ActiveSync mobile devices.
        *   [Scripts/Microsoft 365/MailboxAudit/MailboxAudit.ps1](Scripts/Microsoft%20365/MailboxAudit/MailboxAudit.ps1) - Microsoft 365 Mailbox Sizing, Online Archive & Licensing Auditor GUI.
        *   [Scripts/Microsoft 365/SMSDeprecationCheck/Start-MFACheckerGUI.ps1](Scripts/Microsoft%20365/SMSDeprecationCheck/Start-MFACheckerGUI.ps1) - Audits Entra ID user authentication methods to identify SMS/Voice MFA dependencies.
        *   [Scripts/Microsoft 365/TeamsRooms/AllowForwardOfExternalMeetings/ExternalTeamsRoom.ps1](Scripts/Microsoft%20365/TeamsRooms/AllowForwardOfExternalMeetings/ExternalTeamsRoom.ps1) - Configures Exchange Online calendar processing on Teams Room resource mailboxes to accept external meeting invites.
        *   [Scripts/Microsoft 365/UserLicenceCheck/Console.ps1](Scripts/Microsoft%20365/UserLicenceCheck/Console.ps1) - Exports licensed M365 users alongside last interactive and non-interactive Entra sign-in dates.
    *   **Networking**
        *   [Scripts/Networking/PSDiscovery/Get-SwitchPortInfo.ps1](Scripts/Networking/PSDiscovery/Get-SwitchPortInfo.ps1) - Captures and parses live CDP/LLDP Ethernet frames to identify connected switch name, port, and VLAN.
    *   **Sample**
        *   [Scripts/Sample/SampleScript.ps1](Scripts/Sample/SampleScript.ps1) - Reference script demonstrating repository comment-based help, `Logging.psm1` integration, and error handling.
    *   **Software**
        *   [Scripts/Software/Install-DattoRmmAgent/Install-DattoRmmAgent.ps1](Scripts/Software/Install-DattoRmmAgent/Install-DattoRmmAgent.ps1) - Downloads and silently installs the Datto RMM Agent for a specified site ID.
        *   [Scripts/Software/Install-WindowsUpdatesLTSC/Install-WindowsUpdatesLTSC.ps1](Scripts/Software/Install-WindowsUpdatesLTSC/Install-WindowsUpdatesLTSC.ps1) - Triggers automated Windows Update scan, download, and installation on Windows Enterprise LTSC endpoints.
        *   [Scripts/Software/Update-EdgeBrowser/Update-EdgeBrowser.ps1](Scripts/Software/Update-EdgeBrowser/Update-EdgeBrowser.ps1) - Forces silent update of Microsoft Edge Stable via Enterprise MSI or EdgeUpdate.

## Getting Started

### 1. Creating a New Script
To create a new script with automatic logging enabled, copy the starter template:
```powershell
Copy-Item "Templates/ScriptTemplate.ps1" "Scripts/<Category>/<YourNewTool>/<YourNewTool>.ps1"
```
Open the new script, modify the metadata, and write your logic inside the main `try` block.

### 2. How the Logging Works
Each script that uses the template will:
1. Automatically verify if `C:\Logs` exists. If not, it will be created.
2. Create a log file specific to that script named `[ScriptName]_[YYYY-MM-DD].log` in `C:\Logs`.
3. Use the `Write-ScriptLog` function to write timestamped messages to both the console host and the log file simultaneously.

### 3. Setting execution policy (if blocked)
If you cannot run scripts due to execution policy limits, you can set it in an administrator PowerShell session:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Script Categories & Descriptions

To keep the portal navigation meaningful, the scripts are organized into the following categories under `Scripts/`:
*   **Azure:** Azure VM metadata extraction, Instance Metadata Service (IMDS) queries, cloud resource management, and Azure automation tools.
*   **Clean Up Tools:** Utilities to clean system caches, remove temporary files, purge orphaned installers, and reclaim disk space.
*   **Configuration:** OS parameter customization, power plans, regional UK language settings, and local service configurations.
*   **Datto:** Datto RMM component scripts, privacy mode toggles, policy triggers, and automated status checks.
*   **Intune:** Intune Win32 `.intunewin` app packaging studios, printer deployers, desktop overlays, and MDM policy helpers.
*   **Kiosk:** Dedicated shell locking, UWF management, Windows App LTSC provisioning, and runtime repair scripts for single-app kiosks.
*   **Microsoft 365:** User licensing audits, MFA readiness checks, Exchange quarantine release tools, and Teams Rooms management scripts.
*   **Networking:** CDP/LLDP switch port packet discovery, subnet scanners, DNS resolvers, and network status checks.
*   **Sample:** Reference templates demonstrating standard portal layouts, logging, and error handling.
*   **Software:** Silent software deployment wrappers, Edge browser updaters, LTSC patching, and RMM agent installers.

## Static Code Analysis (Linting)

This repository includes custom settings for **PSScriptAnalyzer** to keep code formatting clean and uniform. If you use VS Code, it will automatically detect [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1) and show warnings/errors inline.

To run linting manually on all files in the repository:
```powershell
Invoke-ScriptAnalyzer -Path . -Settings .\PSScriptAnalyzerSettings.psd1 -Recurse
```
