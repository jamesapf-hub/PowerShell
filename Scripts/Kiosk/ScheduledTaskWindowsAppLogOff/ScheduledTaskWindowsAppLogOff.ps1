# -------------------------------------------------------------------------
# Author: JP
# Date: 06/07/26
# Log Location: C:\Seriun\log\KioskWatchdog.log
# Description: Enhanced monitoring to capture both UWP execution names.
# -------------------------------------------------------------------------

$LogFolder = "C:\Seriun\log"
$LogPath = "$LogFolder\KioskWatchdog.log"

if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }

function Write-KioskLog ($Message) {
    $Timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[$Timestamp] $Message" | Out-File -FilePath $LogPath -Append
}

Write-KioskLog "Kiosk Watchdog v2 Started."

# Wait for either process variations to launch
while ($true) {
    $App = Get-Process -Name "Windows365", "WindowsApp" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($App) {
        Write-KioskLog "Execution detected: $($App.Name) (PID: $($App.Id)). Monitoring session..."
        break
    }
    Start-Sleep -Seconds 2
}

# Wait for the monitored process window to close
$App | Wait-Process
Write-KioskLog "App exit detected. Initiating forceful session shutdown."

# Absolute session logoff command via cmd target override
$SessionId = (Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'").SessionId
if ($SessionId) {
    Start-Process "cmd.exe" -ArgumentList "/c logoff $SessionId" -WindowStyle Hidden
} else {
    # Fallback to standard execution context target
    & logoff
}