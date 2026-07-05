<#
.SYNOPSIS
Cleans leftover Windows upgrade folders (C:\$WINDOWS.~WS and C:\ESD).
.DESCRIPTION
This script checks for the presence of C:\$WINDOWS.~WS and C:\ESD folders, takes ownership if needed, and removes them to recover disk space.
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

function Remove-Folder {
    param([string]$path)

    $testPath = Test-Path $path

    if ($testPath) {
        if ($PSCmdlet.ShouldProcess($path, "Take ownership, grant permissions, and recursively delete folder")) {
            Write-Host "- Taking ownership of '$path'..." -ForegroundColor Cyan
            try {
                takeown /f $path /r /d y | Out-Null
                icacls $path /grant "$($env:USERNAME):(F)" /t | Out-Null
                Write-Host "- Successfully taken ownership!" -ForegroundColor Green
            }
            catch {
                Write-Host "- ERROR: Can't take ownership of '$path'" -ForegroundColor Red
                Write-Host "- Detail: $_" -ForegroundColor Red
                return
            }
            try {
                Write-Host "- Attempting to delete '$path'..." -ForegroundColor Cyan
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-Host "- Successfully removed '$path'" -ForegroundColor Green
            }
            catch {
                Write-Host "- ERROR: Couldn't delete '$path' folder!" -ForegroundColor Red
                Write-Host "- Detail: $_" -ForegroundColor Red
            }             
        }
    }
    else {
        Write-Host "- $path not found" -ForegroundColor Gray
    }
}

function Run-Cleanup {
    Remove-Folder -path 'C:\$WINDOWS.~WS'
    Remove-Folder -path 'C:\ESD'
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
