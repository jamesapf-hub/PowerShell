<#
.SYNOPSIS
    Detection Script for IT Support Desktop Overlay
.DESCRIPTION
    Verifies that C:\ProgramData\ITSupportOverlay\Start-ITOverlay.ps1 and the ITSupportOverlay
    Scheduled Task exist and match the target BuildVersion. Logs results to C:\Logs\DesktopOverlay\Detection.log.
.NOTES
    Script Name   : Detect-ITOverlay.ps1
    Author        : JP
    Version       : 1.1.0
    Log Path      : C:\Logs\DesktopOverlay\Detection.log
#>
param(
    [string]$TargetVersion = "1.0.0"
)

$InstallDir = "C:\ProgramData\ITSupportOverlay"
$ScriptFile = "$InstallDir\Start-ITOverlay.ps1"
$TaskName   = "ITSupportOverlay"
$LogDir     = "C:\Logs\DesktopOverlay"
$LogFile    = "$LogDir\Detection.log"

function Write-DetectLog {
    param([string]$Message)
    try {
        if (-not (Test-Path -Path $LogDir)) {
            $null = New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop
        }
        $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
        Add-Content -Path $LogFile -Value "[$TimeStamp] [DETECTION] $Message" -ErrorAction Stop
    } catch {}
}

$FileExists = Test-Path -Path $ScriptFile
$TaskExists = try { Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop } catch { $null }

if (-not $FileExists -or -not $TaskExists) {
    Write-DetectLog "Status: NOT INSTALLED (File or Task missing)"
    exit 1
}

# Version Check for Overwrite / Update Deployments
$installedContent = try { Get-Content -Path $ScriptFile -Raw -ErrorAction Stop } catch { $null }
if ($installedContent -match '\[string\]\$BuildVersion\s*=\s*"([^"]*)"') {
    $InstalledVersion = $Matches[1]
    Write-DetectLog "Version Check -> Installed: $InstalledVersion | Target: $TargetVersion"
    if ($InstalledVersion -ne $TargetVersion) {
        Write-DetectLog "Status: OUTDATED VERSION ($InstalledVersion != $TargetVersion). Triggering Intune reinstall update."
        exit 1
    }
}

Write-DetectLog "Status: INSTALLED & COMPLIANT ($TargetVersion)"
Write-Output "ITSupportOverlay Installed ($TargetVersion)"
exit 0
