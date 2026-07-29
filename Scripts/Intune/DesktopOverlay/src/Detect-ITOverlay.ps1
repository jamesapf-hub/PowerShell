<#
.SYNOPSIS
    Intune Win32 App Detection Script for IT Support Desktop Overlay
.DESCRIPTION
    Checks if C:\ProgramData\ITSupportOverlay\Start-ITOverlay.ps1 exists
    and if the Scheduled Task ITSupportOverlay is registered.
    Includes optional logging for troubleshooting.
#>

$InstallDir = "C:\ProgramData\ITSupportOverlay"
$ScriptFile = "$InstallDir\Start-ITOverlay.ps1"
$TaskName   = "ITSupportOverlay"
$LogFile    = "$InstallDir\detection.log"

function Write-DetectLog {
    param([string]$Message)
    try {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        Add-Content -Path $LogFile -Value "[$TimeStamp] $Message" -ErrorAction SilentlyContinue
    } catch {}
}

$FileExists = Test-Path -Path $ScriptFile
$TaskExists = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

Write-DetectLog "Detection Check -> ScriptFile Exists: $FileExists | Scheduled Task Exists: $([bool]$TaskExists)"

if ($FileExists -and $TaskExists) {
    Write-DetectLog "Status: INSTALLED"
    Write-Output "ITSupportOverlay Installed"
    exit 0
} else {
    Write-DetectLog "Status: NOT INSTALLED"
    exit 1
}
