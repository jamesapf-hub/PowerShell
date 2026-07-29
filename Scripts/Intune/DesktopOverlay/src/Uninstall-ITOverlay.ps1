<#
.SYNOPSIS
    Uninstaller Script for IT Support Desktop Overlay
.DESCRIPTION
    Stops running overlay processes, deletes the Scheduled Task,
    and removes installation files from C:\ProgramData\ITSupportOverlay.
#>

$InstallDir = "C:\ProgramData\ITSupportOverlay"
$TaskName   = "ITSupportOverlay"

Write-Host "Uninstalling IT Support Desktop Overlay..." -ForegroundColor Cyan

# Unregister Scheduled Task
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false | Out-Null
    Write-Host "Unregistered Scheduled Task: $TaskName" -ForegroundColor Green
}

# Stop running processes running Start-ITOverlay.ps1
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue | 
    Where-Object { $_.CommandLine -like "*Start-ITOverlay.ps1*" } | 
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

# Delete Installation Directory
if (Test-Path -Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Removed installation directory: $InstallDir" -ForegroundColor Green
}

Write-Host "Uninstallation completed successfully." -ForegroundColor Green
