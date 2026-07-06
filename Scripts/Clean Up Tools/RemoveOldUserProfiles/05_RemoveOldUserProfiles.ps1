<#
.SYNOPSIS
    Identifies and removes orphaned AD domain user profiles on a PC.
.DESCRIPTION
    This script scans local user profiles, filters out system, local, and active profiles, and deletes orphaned profiles.
    Saves a persistent log to the Logs directory and runs in WhatIf (dry-run) mode first.
.PARAMETER Force
    Runs the cleanup silently without confirmation prompts.
.PARAMETER ExcludeUsers
    An array of usernames (domain\user) to exclude from the cleanup.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force,
    [string[]]$ExcludeUsers
)

# Define Log Path in UK Date Format (DDMMYY)
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\RemoveOldUserProfiles"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "RemoveOldUserProfiles_$UKDate.log"

# Ensure the log folder exists
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [ConsoleColor]$ForegroundColor = "White"
    )
    $Timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $LogLine = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogLine -ErrorAction SilentlyContinue
    
    $Color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        default   { $ForegroundColor }
    }
    Write-Host $LogLine -ForegroundColor $Color
}

# 1. Gain 'System' level access permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Please run as Administrator!" "ERROR"
    return
}

# Helper to query profiles
function Get-UserProfileList {
    Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-' } |
        ForEach-Object {
            $sid = $_.PSChildName
            $profilePath = $_.ProfilePath
            $username = "Unknown SID ($sid)"
            try {
                $account = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount])
                $username = $account.Value
            } catch {}

            $size = 0
            if ($profilePath -and (Test-Path -Path $profilePath -PathType Container)) {
                try {
                    $size = (Get-ChildItem -LiteralPath $profilePath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                } catch {}
            }

            [PSCustomObject]@{
                SID         = $sid
                ProfilePath = $profilePath
                UserName    = $username
                SizeGB      = [math]::Round($size / 1GB, 3)
            }
        }
}

# 1. Gather profiles
$currentLoggedInUserSID = ""
try {
    $currentLoggedInUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
} catch {}

$wellKnownSids = @()
if ($ExcludeUsers) {
    foreach ($user in $ExcludeUsers) {
        try {
            $wellKnownSids += (New-Object System.Security.Principal.NTAccount($user)).Translate([System.Security.Principal.UIDentifier]).Value
        } catch {}
    }
}

Write-Log "Scanning local user profiles..." "INFO" "Cyan"
$profiles = Get-UserProfileList
$profilesToRemove = @()

foreach ($profile in $profiles) {
    # Skip current logged in user
    if ($profile.SID -eq $currentLoggedInUserSID) { continue }
    # Skip explicitly excluded users
    if ($profile.SID -in $wellKnownSids) { continue }
    
    # Check if old domain account (domain\username where domain != computername and BUILTIN)
    if ($profile.UserName -like "*\*") {
        $parts = $profile.UserName.Split('\')
        $domain = $parts[0]
        if ($domain -ne $env:COMPUTERNAME -and $domain -ne "BUILTIN") {
            $profilesToRemove += $profile
        }
    } else {
        # Unresolved SIDs are likely deleted/orphaned domain accounts
        $profilesToRemove += $profile
    }
}

if ($profilesToRemove.Count -eq 0) {
    Write-Log "No orphaned AD user profiles found." "SUCCESS"
    return
}

Write-Log "Found $($profilesToRemove.Count) orphaned AD profiles." "WARNING"
$profilesToRemove | Format-Table UserName, ProfilePath, SizeGB -AutoSize | Out-String | ForEach-Object {
    if (-not [string]::IsNullOrWhiteSpace($_)) {
        Write-Host $_ -ForegroundColor Yellow
        Add-Content -Path $LogPath -Value $_ -ErrorAction SilentlyContinue
    }
}

function Run-Cleanup {
    foreach ($profile in $profilesToRemove) {
        if ($PSCmdlet.ShouldProcess("$($profile.UserName) ($($profile.ProfilePath))", "Delete User Profile")) {
            try {
                Write-Log "Removing profile for $($profile.UserName)..." "INFO" "Cyan"
                if ($WhatIfPreference) {
                    Write-Host "What if: Performing WMI deletion on profile $($profile.ProfilePath)" -ForegroundColor Gray
                } else {
                    $wmiProfile = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.SID -eq $profile.SID }
                    if ($wmiProfile) {
                        $wmiProfile.Delete()
                        Write-Log "Successfully removed profile for $($profile.UserName)." "SUCCESS"
                    } else {
                        Write-Log "Profile for $($profile.UserName) not found via WMI." "WARNING"
                    }
                }
            } catch {
                Write-Log "Failed to remove profile: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
} else {
    Write-Log "=== STARTING WHATIF DRY-RUN ===" "WARNING"
    $WhatIfPreference = $true
    Run-Cleanup
    
    Write-Host ""
    $confirmation = Read-Host "WhatIf dry-run completed. Do you want to run this for real now? (Y/N)"
    if ($confirmation -eq 'Y' -or $confirmation -eq 'Yes') {
        Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
        $WhatIfPreference = $false
        Run-Cleanup
    } else {
        Write-Log "Cancelled. No changes were made." "INFO" "Gray"
    }
}
