<#
.SYNOPSIS
    Windows Update execution script for LTSC devices using native Windows Update Agent COM objects.
    Ensures zero external module dependencies (does not require PSWindowsUpdate).
    Logs to $env:SystemDrive\Logs\Install-WindowsUpdatesLTSC using UK date format (DDMMYY).
#>

# Define log directory and file name using UK date format (DDMMYY)
$LogDir = "$env:SystemDrive\Logs\Install-WindowsUpdatesLTSC"
$CurrentDate = Get-Date -Format "ddMMyy"
$LogFile = Join-Path $LogDir "WindowsUpdate_$CurrentDate.log"

# Ensure the log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Function Write-Log {
    Param ([string]$Message)
    $Timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Write-Output $LogMessage
    Add-Content -Path $LogFile -Value $LogMessage
}

Write-Log "========================================"
Write-Log "Starting Windows Update script for LTSC (COM Mode)."
Write-Log "========================================"

# Validate Administrator Elevation
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "ERROR: This script must be run as an Administrator (Elevated PowerShell session)."
    Exit 1
}

try {
    Write-Log "Initializing Windows Update Session COM Object..."
    $UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    Write-Log "Searching for applicable missing updates (Software)..."
    # Query only for software updates that are not installed and not hidden
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
    $Count = $SearchResult.Updates.Count

    if ($Count -eq 0) {
        Write-Log "No applicable updates found for this device."
    } else {
        Write-Log "Found $Count update(s) available. Gathering details..."
        
        $UpdatesToDownload = New-Object -ComObject "Microsoft.Update.UpdateColl"
        foreach ($Update in $SearchResult.Updates) {
            Write-Log "  -> [Available] $($Update.Title)"
            
            # Automatically accept EULA if required
            if (-not $Update.EulaAccepted) {
                Write-Log "     Auto-accepting EULA for: $($Update.Title)"
                $Update.AcceptEula()
            }
            $UpdatesToDownload.Add($Update) | Out-Null
        }

        # 1. Download Phase
        Write-Log "Starting download of $Count update(s)..."
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToDownload
        $DownloadResult = $Downloader.Download()
        
        # Verify downloaded updates
        $UpdatesToInstall = New-Object -ComObject "Microsoft.Update.UpdateColl"
        foreach ($Update in $UpdatesToDownload) {
            if ($Update.IsDownloaded) {
                $UpdatesToInstall.Add($Update) | Out-Null
            } else {
                Write-Log "  -> [Skip] Failed to download: $($Update.Title)"
            }
        }

        $InstallCount = $UpdatesToInstall.Count
        if ($InstallCount -eq 0) {
            Write-Log "ERROR: No updates were successfully downloaded. Aborting installation."
            Exit 1
        }

        # 2. Installation Phase
        Write-Log "Starting installation of $InstallCount update(s)..."
        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesToInstall
        
        # Synchronous installation (does not reboot automatically)
        $InstallResult = $Installer.Install()
        
        # Log installation status of each update
        for ($i = 0; $i -lt $UpdatesToInstall.Count; $i++) {
            $Update = $UpdatesToInstall.Item($i)
            $Result = $InstallResult.GetUpdateResult($i)
            
            # ResultCode: 2 = Succeeded, 3 = Succeeded with errors, 4 = Failed, 5 = Aborted
            if ($Result.ResultCode -eq 2) {
                Write-Log "SUCCESS: Installed: $($Update.Title)"
            } elseif ($Result.ResultCode -eq 3) {
                Write-Log "WARNING: Installed with errors: $($Update.Title)"
            } else {
                Write-Log "ERROR: Failed to install: $($Update.Title) (ResultCode: $($Result.ResultCode))"
            }
        }
        
        # 3. Check for Pending Reboot
        # Check both the installation result object and registry status
        $RebootRequired = $InstallResult.RebootRequired -or 
                          (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -or 
                          (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
                          
        if ($RebootRequired) {
            Write-Log "WARNING: A reboot is required to finish installing updates. Please schedule a manual maintenance window."
        } else {
            Write-Log "Updates installed successfully. No reboot required."
        }
    }
}
catch {
    Write-Log "ERROR encountered during the update process: $_"
}

Write-Log "========================================"
Write-Log "Windows Update process completed."
Write-Log "========================================"
