<#
.SYNOPSIS
    Installer Script for IT Support Desktop Overlay
.DESCRIPTION
    Deploys overlay files to C:\ProgramData\ITSupportOverlay, registers a
    Scheduled Task for future user logons, AND immediately launches the overlay
    in the active interactive session (perfect for 'Run in Sandbox' right-click testing).
#>

param(
    [string]$LogPath = "C:\ProgramData\ITSupportOverlay\install.log"
)

# --- Logging Helper ---
function Write-InstallLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")][string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $FormattedMessage = "[$TimeStamp] [$Level] [PID:$PID] $Message"
    
    switch ($Level) {
        "ERROR" { Write-Host $FormattedMessage -ForegroundColor Red }
        "WARN"  { Write-Host $FormattedMessage -ForegroundColor Yellow }
        default { Write-Host $FormattedMessage -ForegroundColor Cyan }
    }
    
    try {
        $LogDir = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path -Path $LogDir)) {
            $null = New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        }
        Add-Content -Path $LogPath -Value $FormattedMessage -ErrorAction SilentlyContinue
    } catch {}
}

Write-InstallLog "--- Starting IT Support Desktop Overlay Installer ---" "INFO"

# --- In-Memory Guard ---
if ($PSScriptRoot -match "^iex" -or [string]::IsNullOrEmpty($PSScriptRoot)) {
    Write-InstallLog "In-memory execution (via iex / WebString) detected. Aborting." "ERROR"
    Write-Error "In-memory execution is not supported. Please extract the package locally before running."
    exit 1
}

$InstallDir = "C:\ProgramData\ITSupportOverlay"
$TaskName   = "ITSupportOverlay"

try {
    Write-InstallLog "Creating installation directory: $InstallDir" "INFO"
    if (-not (Test-Path -Path $InstallDir)) {
        $null = New-Item -Path $InstallDir -ItemType Directory -Force -ErrorAction Stop
        Write-InstallLog "Directory created successfully." "INFO"
    }

    # Copy script files
    $SourceScript = Join-Path -Path $PSScriptRoot -ChildPath "Start-ITOverlay.ps1"
    Write-InstallLog "Looking for source script at: $SourceScript" "INFO"
    if (Test-Path -Path $SourceScript) {
        Copy-Item -Path $SourceScript -Destination "$InstallDir\Start-ITOverlay.ps1" -Force -ErrorAction Stop
        Write-InstallLog "Overlay script copied to $InstallDir\Start-ITOverlay.ps1" "INFO"
    } else {
        Write-InstallLog "Source script Start-ITOverlay.ps1 not found in $PSScriptRoot" "ERROR"
        exit 1
    }

    # Register Scheduled Task (Runs on Logon of ANY user for persistent Intune deployments)
    Write-InstallLog "Registering Scheduled Task: $TaskName" "INFO"

    $Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$InstallDir\Start-ITOverlay.ps1`""
    $Trigger = New-ScheduledTaskTrigger -AtLogOn
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0 -Priority 7

    # Register task under Users group
    try {
        $null = Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -User "Builtin\Users" -Force -ErrorAction Stop
        Write-InstallLog "Scheduled Task '$TaskName' registered successfully." "INFO"
    } catch {
        Write-InstallLog "Scheduled task registration warning (running without task scheduler): $_" "WARN"
    }

    # --- IMMEDIATE LAUNCH FOR ACTIVE SESSION (Crucial for 'Run in Sandbox' right-click testing) ---
    Write-InstallLog "Launching Start-ITOverlay.ps1 directly for immediate desktop display..." "INFO"
    
    # Direct process launch guarantees the overlay shows immediately on Sandbox desktop
    Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$InstallDir\Start-ITOverlay.ps1`"" -WindowStyle Hidden
    Write-InstallLog "Overlay process spawned directly via Start-Process." "INFO"

    Write-InstallLog "Installation completed successfully." "INFO"
} catch {
    Write-InstallLog "Installation failed with error: $_" "ERROR"
    exit 1
}
