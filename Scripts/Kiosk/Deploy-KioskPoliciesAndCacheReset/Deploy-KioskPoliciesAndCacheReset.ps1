# ==============================================================================
# Script Name: Kiosk_NativePolicy_Setup.ps1
# Author:      JP
# Version:     1.3.0
# Description: Deploys native Microsoft Kiosk/Windows App policies via WMS.
#              Manually purges local cached credentials on user logon to avoid
#              mid-session AVD pop-up overlap.
# ==============================================================================

# 1. Path & Metric Targets
$LogFolder     = "$env:SystemDrive\Logs\Deploy-KioskPoliciesAndCacheReset"
$ScriptFolder  = "$env:SystemDrive\Logs\Deploy-KioskPoliciesAndCacheReset\Scripts"
$SetupLogFile  = "$LogFolder\Kiosk_NativePolicy_Setup.txt"
$AppPath       = "HKLM:\SOFTWARE\Microsoft\WindowsApp"
$Win365Path    = "HKLM:\SOFTWARE\Microsoft\Windows365"

# Ensure directories exist
if (-not (Test-Path $LogFolder)) { 
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null 
}
if (-not (Test-Path $ScriptFolder)) { 
    New-Item -ItemType Directory -Path $ScriptFolder -Force | Out-Null 
}

# Standard UK format timestamped logger
function Write-SetupLog ($Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $SetupLogFile -Append
}

# --- CONTEXT CHECK ---
# WMS Deployment Stage (SYSTEM Context)
if ($env:USERNAME -eq "SYSTEM") {
    Write-SetupLog "Execution Context: SYSTEM (WMS Deployment Mode)"
    
    $PermanentPath = "$ScriptFolder\Deploy-KioskPoliciesAndCacheReset.ps1"
    
    try {
        # Copy this script to the permanent local scripts folder
        Copy-Item -Path $MyInvocation.MyCommand.Path -Destination $PermanentPath -Force -ErrorAction Stop
        Write-SetupLog "WMS DEPLOY: Script payload cached to $PermanentPath"
        
        # 1. Provision target HKLM registry paths (requires SYSTEM/Admin privileges)
        if (-not (Test-Path $AppPath)) { 
            New-Item -Path $AppPath -Force | Out-Null
        }
        if (-not (Test-Path $Win365Path)) { 
            New-Item -Path $Win365Path -Force | Out-Null
        }
        
        # Keep the basic auto-logoff when the app closes entirely, and keep SkipFRE
        Set-ItemProperty -Path $AppPath -Name "AutoLogoffEnable" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $Win365Path -Name "SkipFRE" -Value 1 -Type DWord -Force
        
        # Clean AutoLogoffOnSuccessfulConnect to prevent background app overlap
        if (Get-ItemProperty -Path $AppPath -Name "AutoLogoffOnSuccessfulConnect" -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $AppPath -Name "AutoLogoffOnSuccessfulConnect" -Force
            Write-SetupLog "POLICY REMOVAL: Cleaned AutoLogoffOnSuccessfulConnect from HKLM under SYSTEM."
        }

        # 2. Scrub Legacy Custom Run Hooks (requires SYSTEM/Admin privileges)
        if (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "LegacyAppWatcher" -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "LegacyAppWatcher" -Force
            Write-SetupLog "CLEANUP: Removed legacy 'LegacyAppWatcher' registry run key successfully."
        }
        
        # 3. Inject the HKLM Run Key to trigger the script on actual User Logon
        $RunArgs = "powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$PermanentPath`""
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "KioskLogonSetup" -Value $RunArgs -Force | Out-Null
        Write-SetupLog "WMS DEPLOY: Registered HKLM Logon Run Key successfully."
        
        Write-Host "WMS DEPLOY SUCCESS: Payload staged for user logon." -ForegroundColor Green
        exit 0
    }
    catch {
        Write-SetupLog "CRITICAL WMS DEPLOY ERROR: Staging failed. Reason: $_"
        exit 1
    }
}

# --- USER LOGON EXECUTION PHASE ---
# This part executes only when a user logs into the thin client shell.

# Bypasses execution for administrative/other sessions, targeting only the "User" account
if ($env:USERNAME -ne "User") {
    Write-SetupLog "Execution Context: USER ($env:USERNAME) is not the target kiosk account ('User'). Exiting."
    exit 0
}

Write-SetupLog "Execution Context: USER ($env:USERNAME) - Performing Safe Reset"

# 1. Manual Token and Cache Clearance (Logon Only)
# Wipes local app identity caches for the Windows App package prior to initialization
try {
    $TargetCachePaths = @(
        "$env:LOCALAPPDATA\Packages\MicrosoftCorporationII.Windows365_8wekyb3d8bbwe\LocalState",
        "$env:LOCALAPPDATA\Packages\MicrosoftCorporationII.Windows365_8wekyb3d8bbwe\AC\TokenBroker",
        "$env:LOCALAPPDATA\Microsoft\WindowsApp"
    )

    foreach ($Path in $TargetCachePaths) {
        if (Test-Path $Path) {
            # Fixed typo: changed -Recurper to -Recurse
            Remove-Item -Path "$Path\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-SetupLog "LOGON PURGE: Cleared local cache items from $Path"
        }
    }
} catch {
    Write-SetupLog "WARNING: Cache clearance pass encountered locked items. Reason: $_"
}

Write-SetupLog "=== Native Kiosk Engine Provisioning Finished ==="
exit 0