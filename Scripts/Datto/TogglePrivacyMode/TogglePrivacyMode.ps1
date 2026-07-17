<#
.SYNOPSIS
    Toggles the 'PrivacyMode' setting in the CentraStage CagService user.config file and schedules a service restart.
.DESCRIPTION
    This script performs the following actions:
    1. Reads the 'DisplayVersion' from the CentraStage registry path to build the path to the 'user.config' file.
    2. Modifies the 'user.config' XML file to toggle the 'PrivacyMode' value (True/False).
    3. Creates a self-deleting scheduled task that restarts the 'Datto RMM' service one minute later to apply the changes without stalling deployment tools.
.NOTES
    Author     : JP
    Created    : 2025-07-18
    Version    : 1.3
    Log Path   : $env:SystemDrive\Logs\TogglePrivacyMode\TogglePrivacyMode.log
#>

# --- CONFIGURATION & LOGGING ---
$LogDir = "$env:SystemDrive\Logs\TogglePrivacyMode"
$LogFile = "$LogDir\TogglePrivacyMode.log"

if (!(Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    Add-Content -Path $LogFile -Value $LogEntry
    
    switch ($Level) {
        "INFO"    { Write-Host $LogEntry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        default   { Write-Host $LogEntry }
    }
}

Write-Log "Starting TogglePrivacyMode script." "INFO"

# --- 1. LOCATE CONFIGURATION FILE ---
$key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\CentraStage"
$value = "DisplayVersion"

Write-Log "Attempting to read DisplayVersion from registry..." "INFO"
try {
    $version = (Get-ItemProperty -Path $key -Name $value -ErrorAction Stop).$value
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "DisplayVersion is empty or null."
    }
    Write-Log "Found DisplayVersion: $version" "SUCCESS"
} catch {
    Write-Log "Failed to read CentraStage DisplayVersion from registry." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
    exit
}

$path = "C:\Windows\System32\config\systemprofile\AppData\Local\CentraStage\CagService.exe_Url_nin2uaxj2lsg1o0rsz2amvmcciusvum4\"
$file = "\user.config"
$combo = "$path$version$file"

Write-Log "Checking for user.config at: $combo" "INFO"
if (!(Test-Path $combo)) {
    Write-Log "Could not find user.config file at the expected path." "ERROR"
    exit
}

# --- 2. TOGGLE PRIVACY MODE ---
Write-Log "Loading XML configuration..." "INFO"
try {
    $xml = [xml](Get-Content $combo -ErrorAction Stop)
    $node = $xml.configuration.usersettings."CentraStage.Cag.Core.Settings".setting | Where-Object {$_.Name -eq 'PrivacyMode'}
    
    if ($null -eq $node) {
        throw "PrivacyMode setting not found in XML structure."
    }
    
    $oldValue = $node.value
    if ($oldValue -eq 'False') {
        $node.value = 'True'
    } else {
        $node.value = 'False'
    }
    
    Write-Log "Toggled PrivacyMode from '$oldValue' to '$($node.value)'." "SUCCESS"
    
    $xml.Save($combo)
    Write-Log "Successfully saved updated configuration file." "SUCCESS"
} catch {
    Write-Log "Failed to parse or modify user.config." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
    exit
}

# --- 3. SCHEDULE SERVICE RESTART ---
Write-Log "Setting up scheduled task to restart 'Datto RMM' service..." "INFO"
try {
    $STName = "Restart Cags Service"

    if (Get-ScheduledTask -TaskName $STName -ErrorAction SilentlyContinue) {
        Write-Log "Scheduled task '$STName' already exists. Removing it..." "WARNING"
        Unregister-ScheduledTask -TaskName $STName -Confirm:$false -ErrorAction Stop
    }

    $STAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "Restart-Service -Name 'Datto RMM'"
    $STTrigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(1))
    $STPrincipal = New-ScheduledTaskPrincipal -User "NT AUTHORITY\SYSTEM" -RunLevel Highest
    $STSettings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew

    $STSettings.AllowStartIfOnBatteries = $true
    $STSettings.DontStopIfGoingOnBatteries = $true
    $STSettings.StartWhenAvailable = $true

    Register-ScheduledTask -Action $STAction -Trigger $STTrigger -Settings $STSettings -TaskName $STName -Description "Restarts the Cags service for RMM" -Principal $STPrincipal -ErrorAction Stop | Out-Null
    
    $TargetTask = Get-ScheduledTask -TaskName $STName -ErrorAction Stop
    $TargetTask.Triggers[0].EndBoundary = [DateTime]::Now.AddMinutes(5).ToString("yyyy-MM-dd'T'HH:mm:ss")
    $TargetTask.Settings.AllowHardTerminate = $True
    $TargetTask.Settings.DeleteExpiredTaskAfter = 'PT0S'
    $TargetTask.Settings.ExecutionTimeLimit = 'PT1H'
    $TargetTask.Settings.volatile = $False

    $TargetTask | Set-ScheduledTask -ErrorAction Stop | Out-Null
    
    Write-Log "Scheduled task created successfully. Service will restart in 1 minute." "SUCCESS"
} catch {
    Write-Log "Failed to create the scheduled task for service restart." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
}

Write-Log "Script execution completed." "INFO"
