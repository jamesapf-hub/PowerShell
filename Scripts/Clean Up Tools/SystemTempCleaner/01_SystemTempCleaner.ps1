<#
.SYNOPSIS
    Cleans system temporary files, prefetch cache, and empties the Recycle Bin.
.DESCRIPTION
    This script deletes files in C:\Windows\Temp, AppData Temp, and Prefetch folders.
    Saves a persistent log to the Logs directory and runs in WhatIf (dry-run) mode first.
.PARAMETER Force
    Runs the cleanup silently without confirmation prompts.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force
)

# Define Log Path in UK Date Format (DDMMYY)
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\SystemTempCleaner"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "SystemTempCleaner_$UKDate.log"

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

function Run-Cleanup {
    # 1. System Temp
    $SysTemp = "C:\Windows\Temp"
    if (Test-Path $SysTemp) {
        Write-Log "Processing System Temp ($SysTemp)..." "INFO" "Cyan"
        $files = Get-ChildItem -Path "$SysTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
        $count = 0
        foreach ($file in $files) {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete System Temp File")) {
                try {
                    if (-not $WhatIfPreference) {
                        Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                    }
                    $count++
                } catch {
                    Write-Log "Failed to delete: $($file.FullName) - $($_.Exception.Message)" "WARNING"
                }
            }
        }
        Write-Log "Processed $count system temp items." "SUCCESS"
    }

    # 2. User Temp
    $UserTemp = $env:TEMP
    if (Test-Path $UserTemp) {
        Write-Log "Processing User Temp ($UserTemp)..." "INFO" "Cyan"
        $files = Get-ChildItem -Path "$UserTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
        $count = 0
        foreach ($file in $files) {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete User Temp File")) {
                try {
                    if (-not $WhatIfPreference) {
                        Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                    }
                    $count++
                } catch {
                    Write-Log "Failed to delete: $($file.FullName) - $($_.Exception.Message)" "WARNING"
                }
            }
        }
        Write-Log "Processed $count user temp items." "SUCCESS"
    }

    # 3. Windows Prefetch
    $Prefetch = "C:\Windows\Prefetch"
    if (Test-Path $Prefetch) {
        Write-Log "Processing Prefetch ($Prefetch)..." "INFO" "Cyan"
        $files = Get-ChildItem -Path "$Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue
        $count = 0
        foreach ($file in $files) {
            if ($PSCmdlet.ShouldProcess($file.FullName, "Delete Prefetch File")) {
                try {
                    if (-not $WhatIfPreference) {
                        Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                    }
                    $count++
                } catch {
                    Write-Log "Failed to delete: $($file.FullName) - $($_.Exception.Message)" "WARNING"
                }
            }
        }
        Write-Log "Processed $count prefetch items." "SUCCESS"
    }

    # 4. Recycle Bin
    if ($PSCmdlet.ShouldProcess("Recycle Bin", "Empty Recycle Bin")) {
        Write-Log "Emptying Recycle Bin..." "INFO" "Cyan"
        if (-not $WhatIfPreference) {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Log "Recycle Bin cleared." "SUCCESS"
    }
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
    Write-Log "Cleanup complete." "SUCCESS"
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
        Write-Log "Cleanup complete." "SUCCESS"
    } else {
        Write-Log "Cancelled. No changes were made." "INFO" "Gray"
    }
}
