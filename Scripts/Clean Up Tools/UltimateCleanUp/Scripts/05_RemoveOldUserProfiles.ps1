<#
.SYNOPSIS
Identifies and removes orphaned AD domain user profiles on a PC.
.DESCRIPTION
This script scans local user profiles, filters out system, local, and active profiles, and deletes orphaned profiles.
Natively runs in WhatIf (dry-run) mode first, then prompts to execute for real.
Supports -WhatIf and prompts for confirmation unless -Force is specified.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force,
    [string[]]$ExcludeUsers
)

# 1. Gain 'System' level access permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run as Administrator!" -ForegroundColor Red; return
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

Write-Host "Scanning local user profiles..." -ForegroundColor Cyan
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
    Write-Host "No orphaned AD user profiles found." -ForegroundColor Green
    return
}

Write-Host "`nFound $($profilesToRemove.Count) orphaned AD profiles:" -ForegroundColor Yellow
$profilesToRemove | Format-Table UserName, ProfilePath, SizeGB -AutoSize

function Run-Cleanup {
    foreach ($profile in $profilesToRemove) {
        # WMI delete operation natively doesn't support WhatIf automatic logging,
        # so we write a manual ShouldProcess block here:
        if ($PSCmdlet.ShouldProcess("$($profile.UserName) ($($profile.ProfilePath))", "Delete User Profile")) {
            try {
                Write-Host "Removing profile for $($profile.UserName)..." -ForegroundColor Cyan
                $wmiProfile = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.SID -eq $profile.SID }
                if ($wmiProfile) {
                    $wmiProfile.Delete()
                    Write-Host "[+] Successfully removed profile." -ForegroundColor Green
                }
            } catch {
                Write-Host "[-] Failed to remove profile: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
} else {
    Write-Host "=== STARTING WHATIF DRY-RUN ===" -ForegroundColor Yellow
    $WhatIfPreference = $true
    Run-Cleanup
    
    Write-Host ""
    $confirmation = Read-Host "WhatIf dry-run completed. Do you want to run this for real now? (Y/N)"
    if ($confirmation -eq 'Y' -or $confirmation -eq 'Yes') {
        Write-Host "`n=== RUNNING REAL CLEANUP ===" -ForegroundColor Green
        $WhatIfPreference = $false
        Run-Cleanup
    } else {
        Write-Host "`nCancelled. No changes were made." -ForegroundColor Gray
    }
}
