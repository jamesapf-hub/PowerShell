<#
.SYNOPSIS
Cleans browser cache files.
.DESCRIPTION
This script simulates cleaning Google Chrome and Microsoft Edge cache files.
#>

[CmdletBinding()]
param(
    [switch]$Force
)

Write-Host "[*] Starting Browser Cache Cleanup..." -ForegroundColor Cyan
Start-Sleep -Seconds 3
Write-Host "[+] Cleaning Chrome cache files... Done." -ForegroundColor Green
Write-Host "[+] Cleaning Edge cache files... Done." -ForegroundColor Green
Write-Host "[+] Browser Cache Cleanup complete." -ForegroundColor Green
