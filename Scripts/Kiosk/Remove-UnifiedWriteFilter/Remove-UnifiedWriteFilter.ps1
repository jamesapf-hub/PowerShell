# Define paths
$LogFolder = "C:\Seriun\log"
$LogFile   = "$LogFolder\UWFdisable.txt"
$TargetDir = "C:\Program Files\DACC"
$AppExe    = "daccnotifier.exe"

# Ensure log directory exists
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "ddMMyy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
}

Write-Log "=== Script Started ==="

# 1. Handle File Renaming
if (Test-Path "$TargetDir\$AppExe") {
    try {
        Rename-Item -Path "$TargetDir\$AppExe" -NewName "$AppExe.bak" -Force -ErrorAction Stop
        Write-Log "SUCCESS: Renamed $AppExe to $AppExe.bak"
    } catch {
        Write-Log "ERROR: Failed to rename file. Reason: $_"
    }
} else {
    Write-Log "INFO: $AppExe not found or already renamed. Skipping."
}

# 2. Completely Purge Unified Write Filter via PowerShell
Write-Log "Checking Unified Write Filter (UWF) status..."

$UwfEnabled = $false
if (Get-Command uwfmgr -ErrorAction SilentlyContinue) {
    try {
        $Config = uwfmgr filter get-config
        if ($Config -match "Filter state: ON" -or $Config -match "Unified Write Filter is enabled") {
            $UwfEnabled = $true
        }
    } catch {
        Write-Log "WARNING: Could not query UWF status via uwfmgr. Reason: $_"
    }
}

if ($UwfEnabled) {
    Write-Log "WARNING: Unified Write Filter is currently ENABLED. Disabling it now..."
    try {
        & uwfmgr filter disable | Out-String | Write-Log
        Write-Log "SUCCESS: UWF filter set to disabled. Initiating system reboot to apply changes and prevent BSOD..."
        
        # Force reboot to load OS without the filter active
        Start-Sleep -Seconds 5
        Restart-Computer -Force -Confirm:$false
        exit 0
    } catch {
        Write-Log "ERROR: Failed to disable UWF filter. Reason: $_"
        exit 1
    }
} else {
    Write-Log "INFO: UWF is disabled (or not present). Safe to purge Client-UnifiedWriteFilter feature."
    try {
        # -Remove completely uninstalls/purges the feature payload from the OS side-by-side store
        Disable-WindowsOptionalFeature -Online -FeatureName "Client-UnifiedWriteFilter" -Remove -NoRestart -ErrorAction Stop
        Write-Log "SUCCESS: Unified Write Filter completely uninstalled and purged."
    } catch {
        Write-Log "ERROR: Failed to purge UWF feature. Reason: $_"
    }
}

Write-Log "=== Script Finished ==="
exit 0