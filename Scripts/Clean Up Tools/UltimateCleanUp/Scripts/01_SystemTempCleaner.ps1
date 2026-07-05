<#
.SYNOPSIS
Cleans system temporary files, prefetch cache, and empties the Recycle Bin.
.DESCRIPTION
This script deletes files in C:\Windows\Temp, AppData Temp, and Prefetch folders.
Natively runs in WhatIf (dry-run) mode first, then prompts to execute for real.
Supports -WhatIf and prompts for confirmation unless -Force is specified.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force
)

# 1. Gain 'System' level access permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run as Administrator!" -ForegroundColor Red; return
}

function Run-Cleanup {
    # 1. System Temp
    $SysTemp = "C:\Windows\Temp"
    if (Test-Path $SysTemp) {
        Write-Host "Processing System Temp..." -ForegroundColor Cyan
        $files = Get-ChildItem -Path "$SysTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
            } catch {}
        }
    }

    # 2. User Temp
    $UserTemp = $env:TEMP
    if (Test-Path $UserTemp) {
        Write-Host "Processing User Temp..." -ForegroundColor Cyan
        $files = Get-ChildItem -Path "$UserTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
            } catch {}
        }
    }

    # 3. Windows Prefetch
    $Prefetch = "C:\Windows\Prefetch"
    if (Test-Path $Prefetch) {
        Write-Host "Processing Prefetch..." -ForegroundColor Cyan
        $files = Get-ChildItem -Path "$Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                Remove-Item -Path $file.FullName -Recurse -Force -ErrorAction Stop
            } catch {}
        }
    }

    # 4. Recycle Bin
    Write-Host "Processing Recycle Bin..." -ForegroundColor Cyan
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
}

# Execution Flow
if ($Force) {
    $WhatIfPreference = $false
    Run-Cleanup
    Write-Host "[+] Cleanup complete." -ForegroundColor Green
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
        Write-Host "[+] Cleanup complete." -ForegroundColor Green
    } else {
        Write-Host "`nCancelled. No changes were made." -ForegroundColor Gray
    }
}
