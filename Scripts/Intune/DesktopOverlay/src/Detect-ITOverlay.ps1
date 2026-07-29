param(
    [string]$TargetVersion = "1.0.0"
)

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

if (-not $FileExists -or -not $TaskExists) {
    Write-DetectLog "Status: NOT INSTALLED (File or Task missing)"
    exit 1
}

# Version Check for Overwrite / Update Deployments
$installedContent = Get-Content -Path $ScriptFile -Raw -ErrorAction SilentlyContinue
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
