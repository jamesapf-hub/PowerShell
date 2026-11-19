<#
.SYNOPSIS
    Uninstaller Script for IT Support Desktop Overlay
.DESCRIPTION
    Stops running overlay processes, deletes the Scheduled Task,
    and removes installation files from C:\ProgramData\ITSupportOverlay.
.NOTES
    Script Name   : Uninstall-ITOverlay.ps1
    Author        : JP
    Version       : 1.1.0
    Log Path      : C:\Logs\DesktopOverlay\Uninstall.log
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$LogPath = "C:\Logs\DesktopOverlay\Uninstall.log"
)

$InstallDir = "C:\ProgramData\ITSupportOverlay"
$TaskName   = "ITSupportOverlay"

function Write-UninstallLog {
    param([string]$Message, [string]$Level = "INFO")
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $Entry = "[$TimeStamp] [$Level] $Message"
    Write-Output $Entry
    try {
        $LogDir = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path -Path $LogDir)) {
            $null = New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop
        }
        Add-Content -Path $LogPath -Value $Entry -ErrorAction Stop
    } catch {}
}

Write-UninstallLog -Message "=== Uninstalling IT Support Desktop Overlay ===" -Level "INFO"

# Unregister Scheduled Task
$existingTask = try { Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop } catch { $null }
if ($existingTask) {
    if ($PSCmdlet.ShouldProcess($TaskName, "Unregister Scheduled Task")) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false | Out-Null
        Write-UninstallLog -Message "Unregistered Scheduled Task: $TaskName" -Level "SUCCESS"
    }
}

# Stop running processes running Start-ITOverlay.ps1
$overlayProcs = try { Get-CimInstance -ClassName Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop } catch { @() }
$overlayProcs | 
    Where-Object { $_.CommandLine -like "*Start-ITOverlay.ps1*" } | 
    ForEach-Object {
        if ($PSCmdlet.ShouldProcess("PID $($_.ProcessId)", "Stop running Start-ITOverlay.ps1 process")) {
            try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {}
        }
    }

# Delete Installation Directory
if (Test-Path -Path $InstallDir) {
    if ($PSCmdlet.ShouldProcess($InstallDir, "Remove installation directory")) {
        try { Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction Stop } catch {}
        Write-UninstallLog -Message "Removed installation directory: $InstallDir" -Level "SUCCESS"
    }
}

Write-UninstallLog -Message "=== Uninstallation completed successfully ===" -Level "SUCCESS"
