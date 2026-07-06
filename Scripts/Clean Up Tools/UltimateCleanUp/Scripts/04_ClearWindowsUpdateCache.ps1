<#
.SYNOPSIS
Clears the Windows Update SoftwareDistribution Download cache.
.DESCRIPTION
This script stops Windows Update services (wuauserv, bits), purges the Download folder contents, and restarts the services.
Natively runs in WhatIf (dry-run) mode first, then prompts to execute for real.
Supports -WhatIf and prompts for confirmation unless -Force is specified.
.NOTES
    Date Format: UK (DDMMYY)
    Log Path   : $env:SystemDrive\Logs\UltimateCleanUp\SD_Clear_DDMMYY.log
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force
)

# 1. Environment & Logging Setup
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\UltimateCleanUp"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "SD_Clear_$UKDate.log"

# Ensure the log folder exists
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $LogLine = "[$Timestamp] [$Level] $Message"
    
    # Write to console and append to file
    Write-Host $LogLine
    Add-Content -Path $LogPath -Value $LogLine
}

# 2. Administrator Elevation Check
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "Please run as Administrator!" -ForegroundColor Red; return
}

# Scan cache first
$TargetFolder = "C:\Windows\SoftwareDistribution\Download"
$LiteralTarget = "\\?\$TargetFolder"
$filesCount = 0
$totalSize = 0

if (Test-Path -LiteralPath $LiteralTarget) {
    Write-Host "Scanning Windows Update cache... Please wait." -ForegroundColor Cyan
    $stats = Get-ChildItem -Path $TargetFolder -Recurse -File -Force -ErrorAction SilentlyContinue | 
             Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
    if ($stats) {
        $filesCount = $stats.Count
        $totalSize = $stats.Sum
    }
}
$sizeMB = [math]::Round($totalSize / 1MB, 2)

Write-Host "`nFound $filesCount files in Windows Update cache ($sizeMB MB)." -ForegroundColor Yellow

if ($filesCount -eq 0) {
    Write-Host "Windows Update cache is already empty." -ForegroundColor Green
    return
}

function Run-Cleanup {
    Write-Log "Starting SoftwareDistribution Cleanup..."
    
    # Stop services
    $Services = @("wuauserv", "bits")
    foreach ($Service in $Services) {
        $ServiceStatus = Get-Service -Name $Service -ErrorAction SilentlyContinue
        if ($ServiceStatus -and $ServiceStatus.Status -eq 'Running') {
            Write-Log "Stopping service: $Service"
            # Stop-Service natively respects WhatIf
            Stop-Service -Name $Service -Force
            Start-Sleep -Seconds 1
        }
    }
    
    # Clear folder
    if (Test-Path -LiteralPath $LiteralTarget) {
        try {
            Write-Log "Purging items from: $TargetFolder"
            Get-ChildItem -LiteralPath $LiteralTarget -Force | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
            }
            Write-Log "Download folder successfully cleared." "SUCCESS"
        }
        catch {
            Write-Log "Failed to clear natively: $($_.Exception.Message). Fallback to Robocopy sync..." "WARNING"
            if ($WhatIfPreference) {
                Write-Host "What if: Performing Robocopy purge on target $TargetFolder"
            } else {
                # Bulletproof Fallback: Robocopy an empty temp dir into it
                $TempEmpty = "C:\Windows\Temp\Empty_$UKDate"
                New-Item -Path $TempEmpty -ItemType Directory -Force | Out-Null
                & robocopy $TempEmpty $TargetFolder /MIR /R:0 /W:0 /NJH /NJS /NDL /NC /NS | Out-Null
                Remove-Item -Path $TempEmpty -Force -ErrorAction SilentlyContinue
                Write-Log "Robocopy fallback purge completed." "SUCCESS"
            }
        }
    }
    
    # Start services
    foreach ($Service in $Services) {
        Write-Log "Restarting service: $Service"
        Start-Service -Name $Service
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
