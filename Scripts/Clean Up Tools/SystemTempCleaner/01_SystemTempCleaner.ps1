<#
.SYNOPSIS
    Cleans system temporary files, prefetch cache, and empties the Recycle Bin.
.DESCRIPTION
    This script deletes files in C:\Windows\Temp, AppData Temp, and Prefetch folders.
    Saves a persistent log to the Logs directory and runs in WhatIf (dry-run) mode first.
.PARAMETER Force
    Runs the cleanup silently without confirmation prompts.
#>

[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$DryRun
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
            try {
                if ($WhatIfPreference) {
                    Write-Host "What if: Performing operation 'Delete System Temp File' on Target '$($file.FullName)'" -ForegroundColor Gray
                } else {
                    Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                }
                $count++
            } catch {
                # Silently skip locked files to prevent console spam
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
            try {
                if ($WhatIfPreference) {
                    Write-Host "What if: Performing operation 'Delete User Temp File' on Target '$($file.FullName)'" -ForegroundColor Gray
                } else {
                    Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                }
                $count++
            } catch {
                # Silently skip locked files to prevent console spam
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
            try {
                if ($WhatIfPreference) {
                    Write-Host "What if: Performing operation 'Delete Prefetch File' on Target '$($file.FullName)'" -ForegroundColor Gray
                } else {
                    Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
                }
                $count++
            } catch {
                # Silently skip locked files to prevent console spam
            }
        }
        Write-Log "Processed $count prefetch items." "SUCCESS"
    }

    # 4. Recycle Bin
    Write-Log "Emptying Recycle Bin..." "INFO" "Cyan"
    if ($WhatIfPreference) {
        Write-Host "What if: Performing operation 'Empty Recycle Bin' on Target 'Recycle Bin'" -ForegroundColor Gray
    } else {
        try {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Recycle Bin cleared." "SUCCESS"
        } catch {}
    }
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
    Write-Log "Cleanup complete." "SUCCESS"
} elseif ($DryRun) {
    Write-Log "=== STARTING WHATIF DRY-RUN (Strict Mode) ===" "WARNING"
    $WhatIfPreference = $true
    Run-Cleanup
    Write-Log "Strict DryRun completed. No changes were made." "INFO" "Gray"
} else {
    if (-not [Environment]::UserInteractive) {
        Write-Log "Non-interactive session detected. Forcing Strict DryRun to prevent hanging." "WARNING" "Yellow"
        $WhatIfPreference = $true
        Run-Cleanup
        Write-Log "DryRun complete. Re-run with -Force to execute actual cleanup." "INFO" "Gray"
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
}
