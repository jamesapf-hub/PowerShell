# Define log paths and targets
$LogFolder = "$env:SystemDrive\Logs\Schedule-KioskAppRuntimeFix"
$LogFile   = "$LogFolder\Kiosk_Runtime_Fix.txt"
$TaskName  = "Kiosk-ShellAppRuntime-Fix"
$RegPath   = "HKLM\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"

# 1. Create Log Folder if it doesn't exist
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# Logging function to format timestamps
function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
}

Write-Log "=== Kiosk Master Setup with App Reset Started ==="

# --- STEP A: APPLY REGISTRY FIX ---
Write-Log "Applying Custom Shell UWP registry fix..."
try {
    if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
    Set-ItemProperty -Path $RegPath -Name "EnableUWPAppsForCustomShell" -Value 1 -Type DWord -ErrorAction Stop
    Write-Log "SUCCESS: Registry key set to 1 under $RegPath"
} catch {
    Write-Log "ERROR: Failed to update registry key. Reason: $_"
}

# --- STEP B: BUILD COMBINED LOGON TASK ---
Write-Log "Building Scheduled Task with App Reset arguments..."
try {
    # This command block runs hidden at logon to completely reset the Windows App and start the runtime.
    # Note: *Microsoft.RemoteDesktop* is the inner system name for the Windows App package.
    $ActionScript = 'PowerShell.exe -NoProfile -WindowStyle Hidden -Command "' +
        '$App = Get-AppxPackage -Name *Microsoft.RemoteDesktop* -AllUsers;' +
        'if ($App) { Reset-AppxPackage -Package $App.PackageFullName -ErrorAction SilentlyContinue };' +
        'start C:\Windows\System32\ShellAppRuntime.exe' +
    '"'

    $Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c $ActionScript"
    $Trigger = New-ScheduledTaskTrigger -AtLogOn
    $Principal = New-ScheduledTaskPrincipal -GroupId "INTERACTIVE" -RunLevel Highest
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Compatibility Win8

    # Remove old version if it exists to prevent conflicts
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
    }

    # Register the new task
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Description "Resets Windows App local data and starts background UWP infrastructure on login." -ErrorAction Stop
    
    Write-Log "SUCCESS: Scheduled Task '$TaskName' created successfully."
} catch {
    Write-Log "ERROR: Failed to create Scheduled Task. Reason: $_"
}

Write-Log "=== Kiosk Master Setup with App Reset Finished ==="
exit 0