<#
.SYNOPSIS
    Cleans orphan Adobe installer patch files (.msp) to recover disk space.
.DESCRIPTION
    This script scans C:\Windows\Installer for Adobe installer patch files, calculates the potential savings, and deletes them.
    Natively runs in WhatIf (dry-run) mode first, then prompts to execute for real.
    Supports -WhatIf and prompts for confirmation unless -Force is specified.
.NOTES
    Date Format: UK (DDMMYY)
    Log Path   : $env:SystemDrive\Logs\CleanAdobeOrphanInstallers\CleanAdobeOrphanInstallers_DDMMYY.log
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force
)

# 1. Gain 'System' level access permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run as Administrator!" -ForegroundColor Red; return
}

# 2. Environment & Logging Setup
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\CleanAdobeOrphanInstallers"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "CleanAdobeOrphanInstallers_$UKDate.log"

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
    Write-Host $Message -ForegroundColor $Color
}

Write-Log "=== Clean Adobe Orphan Installers Started ==="

$installer = New-Object -ComObject WindowsInstaller.Installer
$adobeFiles = @()
$totalAdobeSizeGB = 0

# Stop the installer service to unlock files
Stop-Service -Name "msiserver" -ErrorAction SilentlyContinue
Write-Log "Scanning C:\Windows\Installer for Adobe files... Please wait." -ForegroundColor Cyan
$files = Get-ChildItem "C:\Windows\Installer\*.msp"

# --- PHASE 1: SCANNING ---
foreach ($file in $files) {
    try {
        $sumInfo = $installer.GetType().InvokeMember("SummaryInformation", "GetProperty", $null, $installer, @([string]$file.FullName, 0))
        $author = $sumInfo.GetType().InvokeMember("Property", "GetProperty", $null, $sumInfo, @(8))
        $subject = $sumInfo.GetType().InvokeMember("Property", "GetProperty", $null, $sumInfo, @(3))
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sumInfo) | Out-Null
        
        # Search for Adobe's "Stealth" signatures
        if ($author -like "*Adobe*" -or $author -like "*TGT_*ToUPG*" -or $subject -like "*Acrobat*" -or $subject -like "*Adobe*") {
            $sizeGB = $file.Length/1GB
            $adobeFiles += [PSCustomObject]@{
                FullName = $file.FullName
                Name     = $file.Name
                SizeGB   = $sizeGB
            }
            $totalAdobeSizeGB += $sizeGB
        }
    } catch {}
}

$roundedTotal = [math]::Round($totalAdobeSizeGB, 2)
if ($adobeFiles.Count -eq 0) {
    Write-Log "No Adobe orphan files found." -Level "SUCCESS"
    Start-Service -Name "msiserver" -ErrorAction SilentlyContinue
    return
}

Write-Log "Found $($adobeFiles.Count) Adobe files totaling: $roundedTotal GB" -Level "WARNING"

function Run-Cleanup {
    $totalRecovered = 0
    foreach ($file in $adobeFiles) {
        try {
            # Take ownership and grant permissions
            takeown /f $file.FullName /a > $null
            icacls $file.FullName /grant *S-1-5-32-544:F > $null
            
            # Remove the file (Remove-Item natively respects WhatIf)
            Remove-Item $file.FullName -Force -ErrorAction Stop
            $totalRecovered += $file.SizeGB
        } catch {
            Write-Log "Could not delete $($file.Name)." -Level "ERROR"
        }
    }
    if (-not $WhatIfPreference) {
        $finalRecovered = [math]::Round($totalRecovered, 2)
        Write-Log "Success! Total space recovered: $finalRecovered GB" -Level "SUCCESS"
    }
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
} else {
    Write-Log "=== STARTING WHATIF DRY-RUN ===" -Level "WARNING"
    $WhatIfPreference = $true
    Run-Cleanup
    
    Write-Log ""
    $confirmation = Read-Host "WhatIf dry-run completed. Do you want to run this for real now? (Y/N)"
    if ($confirmation -eq 'Y' -or $confirmation -eq 'Yes') {
        Write-Log "=== RUNNING REAL CLEANUP ===" -Level "SUCCESS"
        $WhatIfPreference = $false
        Run-Cleanup
    } else {
        Write-Log "Cancelled. No changes were made." -ForegroundColor Gray
    }
}

Start-Service -Name "msiserver" -ErrorAction SilentlyContinue
Write-Log "=== Clean Adobe Orphan Installers Finished ==="
