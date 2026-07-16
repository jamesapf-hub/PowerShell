<#
.SYNOPSIS
    Cleans leftover Windows upgrade folders (C:\$WINDOWS.~WS and C:\ESD).
.DESCRIPTION
    This script checks for the presence of C:\$WINDOWS.~WS and C:\ESD folders, takes ownership if needed, and removes them to recover disk space.
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
$LogDirectory = "$env:SystemDrive\Logs\TidyUpWindows11"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "TidyUpWindows11_$UKDate.log"

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

function Remove-Folder {
    param([string]$path)

    $testPath = Test-Path $path

    if ($testPath) {
        $folderSize = 0
        try {
            $fso = New-Object -ComObject Scripting.FileSystemObject
            $fsoSize = $fso.GetFolder($path).Size
            if ($null -ne $fsoSize) { $folderSize = $fsoSize }
        } catch {}

        if ($folderSize -eq 0) {
            try {
                $fallbackSize = (Get-ChildItem -LiteralPath $path -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($null -ne $fallbackSize) { $folderSize = $fallbackSize }
            } catch {}
        }
        $script:totalBytesSaved += $folderSize

        Write-Log "Taking ownership of '$path'..." "INFO" "Cyan"
        if ($WhatIfPreference) {
            Write-Host "What if: takeown /f $path /r /d y" -ForegroundColor Gray
            Write-Host "What if: icacls $path /grant \"\$($env:USERNAME):(F)\" /t" -ForegroundColor Gray
            Write-Host "What if: Remove-Item -Path $path -Recurse -Force" -ForegroundColor Gray
        } else {
            try {
                takeown /f $path /r /d y | Out-Null
                icacls $path /grant "$($env:USERNAME):(F)" /t | Out-Null
                Write-Log "Successfully taken ownership of '$path'!" "SUCCESS"
            }
            catch {
                Write-Log "ERROR: Can't take ownership of '$path'. Detail: $_" "ERROR"
                return
            }
            try {
                Write-Log "Attempting to delete '$path'..." "INFO" "Cyan"
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-Log "Successfully removed '$path'" "SUCCESS"
            }
            catch {
                Write-Log "ERROR: Couldn't delete '$path' folder! Detail: $_" "ERROR"
            }
        }
    }
    else {
        Write-Log "$path not found" "INFO" "Gray"
    }
}

function Run-Cleanup {
    $script:totalBytesSaved = 0

    Remove-Folder -path 'C:\$WINDOWS.~WS'
    Remove-Folder -path 'C:\ESD'

    $savedMB = [math]::Round($script:totalBytesSaved / 1MB, 2)
    Write-Log "Total space processed: $savedMB MB" "SUCCESS" "Green"
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
