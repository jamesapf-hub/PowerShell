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
    [switch]$DryRun,
    [switch]$ConsoleMode,
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
    $fso = New-Object -ComObject Scripting.FileSystemObject
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    
    Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-' } |
        ForEach-Object {
            $sid = $_.PSChildName
            $profilePath = $_.GetValue("ProfileImagePath")
            if (-not [string]::IsNullOrWhiteSpace($profilePath)) {
                $folderName = Split-Path $profilePath -Leaf
                $username = "Unknown SID ($folderName)"
            } else {
                $username = "Unknown SID ($sid) [Path Missing]"
            }
            
            try {
                $account = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount])
                $username = $account.Value
            } catch {}

            $size = 0
            if ($profilePath -and (Test-Path -Path $profilePath -PathType Container)) {
                try {
                    $size = $fso.GetFolder($profilePath).Size
                } catch {
                    # Fallback to standard Get-ChildItem if COM fails
                    try { $size = (Get-ChildItem -LiteralPath $profilePath -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum } catch {}
                }
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
                    Write-Host "What if: Performing cleanup on profile $($profile.ProfilePath) (SID: $($profile.SID))" -ForegroundColor Gray
                } else {
                    $deleted = $false
                    $wmiProfile = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.SID -eq $profile.SID }
                    if ($wmiProfile) {
                        $wmiProfile.Delete()
                        Write-Log "Successfully removed profile via WMI." "SUCCESS"
                        $deleted = $true
                    } else {
                        Write-Log "Profile not found via WMI. Attempting forced manual cleanup..." "WARNING"
                    }

                    # Force Cleanup Fallback for Ghost SIDs
                    if (-not $deleted) {
                        # 1. Delete physical folder if it exists
                        if ($profile.ProfilePath -and (Test-Path -Path $profile.ProfilePath)) {
                            Remove-Item -Path $profile.ProfilePath -Force -Recurse -ErrorAction SilentlyContinue
                            Write-Log "Force deleted stranded profile folder." "SUCCESS"
                        }
                        
                        # 2. Delete the dead registry key
                        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($profile.SID)"
                        if (Test-Path -Path $regPath) {
                            Remove-Item -Path $regPath -Force -Recurse -ErrorAction SilentlyContinue
                            Write-Log "Successfully purged dead registry key." "SUCCESS"
                        }
                    }
                }
            } catch {
                Write-Log "Failed to remove profile: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# Execution Flow
if (-not [Environment]::UserInteractive) {
    Write-Log "Non-interactive session detected. Forcing Console Mode." "WARNING" "Yellow"
    $ConsoleMode = $true
}

if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
} elseif ($DryRun) {
    Write-Log "=== STARTING WHATIF DRY-RUN (Strict Mode) ===" "WARNING"
    $WhatIfPreference = $true
    Run-Cleanup
    Write-Log "Strict DryRun completed. No changes were made." "INFO" "Gray"
} elseif ($ConsoleMode) {
    Write-Log "Launching interactive console selection menu..." "INFO" "Cyan"
    Write-Host "`nAll profiles below will be permanently deleted. Select the ones you want to delete:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $profilesToRemove.Count; $i++) {
        $profile = $profilesToRemove[$i]
        Write-Host "[$i] $($profile.UserName) | Size: $($profile.SizeGB) GB"
    }
    
    Write-Host ""
    $selection = Read-Host "Enter numbers to DELETE separated by commas (e.g. 0,2,3), or 'all' to delete everything, or 'none' to cancel"
    
    if ($selection -eq 'none' -or [string]::IsNullOrWhiteSpace($selection)) {
        Write-Log "Cancelled by user. No changes were made." "INFO" "Gray"
    } elseif ($selection -eq 'all') {
        Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
        $WhatIfPreference = $false
        Run-Cleanup
    } else {
        $approvedProfiles = @()
        $indexes = $selection -split ',' | ForEach-Object { $_.Trim() }
        foreach ($index in $indexes) {
            if ($index -match '^\d+$') {
                $idx = [int]$index
                if ($idx -ge 0 -and $idx -lt $profilesToRemove.Count) {
                    $approvedProfiles += $profilesToRemove[$idx]
                }
            }
        }
        
        $profilesToRemove = $approvedProfiles
        if ($profilesToRemove.Count -eq 0) {
            Write-Log "No valid profiles selected. No changes were made." "INFO" "Gray"
        } else {
            Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
            $WhatIfPreference = $false
            Run-Cleanup
        }
    }
} else {
    Write-Log "Launching interactive profile selection menu..." "INFO" "Cyan"
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select Profiles to Delete'
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(460,20)
    $label.Text = 'All profiles below will be permanently deleted. UNTICK any you want to KEEP:'
    $form.Controls.Add($label)
    
    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10,40)
    $checkedListBox.Size = New-Object System.Drawing.Size(460,260)
    $checkedListBox.CheckOnClick = $true
    
    foreach ($profile in $profilesToRemove) {
        $item = "$($profile.UserName) | Size: $($profile.SizeGB) GB"
        $checkedListBox.Items.Add($item, $true) | Out-Null
    }
    $form.Controls.Add($checkedListBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150,310)
    $okButton.Size = New-Object System.Drawing.Size(90,30)
    $okButton.Text = 'Delete Selected'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(260,310)
    $cancelButton.Size = New-Object System.Drawing.Size(75,30)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $approvedProfiles = @()
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            if ($checkedListBox.GetItemChecked($i)) {
                $approvedProfiles += $profilesToRemove[$i]
            }
        }
        
        $profilesToRemove = $approvedProfiles
        if ($profilesToRemove.Count -eq 0) {
            Write-Log "All profiles were unchecked. No changes were made." "INFO" "Gray"
        } else {
            Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
            $WhatIfPreference = $false
            Run-Cleanup
        }
    } else {
        Write-Log "Cancelled by user. No changes were made." "INFO" "Gray"
    }
}
