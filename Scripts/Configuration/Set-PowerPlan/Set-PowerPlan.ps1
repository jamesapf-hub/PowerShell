# Standard UK format timestamped logger
$LogFolder = "$env:SystemDrive\Logs\Set-PowerPlan"
$LogFile   = "$LogFolder\PowerPlanConfig.txt"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
    Write-Output "[ $TimeStamp ] $Message"
}

Write-Log "=== Power Configuration Script Started ==="

# Get active power scheme name
try {
    $ActiveSchemeOutput = powercfg /getactivescheme
    Write-Log "Current Active Power Scheme: $ActiveSchemeOutput"
} catch {
    Write-Log "WARNING: Could not retrieve active power scheme details."
}

# Configure timeouts to 0 (Never) for both AC (plugged in) and DC (battery)
$Configurations = @(
    @("monitor-timeout-ac", "0", "AC Display Timeout"),
    @("monitor-timeout-dc", "0", "DC Display Timeout"),
    @("standby-timeout-ac", "0", "AC Sleep/Standby Timeout"),
    @("standby-timeout-dc", "0", "DC Sleep/Standby Timeout"),
    @("hibernate-timeout-ac", "0", "AC Hibernate Timeout"),
    @("hibernate-timeout-dc", "0", "DC Hibernate Timeout"),
    @("disk-timeout-ac", "0", "AC Disk Sleep Timeout"),
    @("disk-timeout-dc", "0", "DC Disk Sleep Timeout")
)

foreach ($config in $Configurations) {
    $setting = $config[0]
    $value = $config[1]
    $description = $config[2]
    
    try {
        powercfg /change $setting $value
        Write-Log "SUCCESS: Set $description to $value (Never)"
    } catch {
        Write-Log "ERROR: Failed to set $description to $value. Reason: $_"
    }
}

# Disable hibernation (frees up disk space and disables Fast Startup to ensure 100% clean boots)
try {
    powercfg /h off
    Write-Log "SUCCESS: Hibernation disabled (which also disables Fast Startup for clean boots)"
} catch {
    Write-Log "WARNING: Could not disable hibernation. Reason: $_"
}

Write-Log "=== Power Configuration Script Finished ==="
exit 0
