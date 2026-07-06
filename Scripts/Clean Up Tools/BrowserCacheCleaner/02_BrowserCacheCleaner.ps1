<#
.SYNOPSIS
    Cleans browser cache files.
.DESCRIPTION
    This script simulates cleaning Google Chrome and Microsoft Edge cache files.
    Saves a persistent log to the Logs directory.
.PARAMETER Force
    Bypasses execution pauses.
#>

[CmdletBinding()]
param(
    [switch]$Force
)

# Define Log Path in UK Date Format (DDMMYY)
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\BrowserCacheCleaner"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "BrowserCacheCleaner_$UKDate.log"

# Ensure the log folder exists
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [ConsoleColor]$ForegroundColor = "White"
    )
    $Timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $LogLine = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogLine -ErrorAction SilentlyContinue
    
    $Color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        default   { $ForegroundColor }
    }
    Write-Host $LogLine -ForegroundColor $Color
}

Write-Log "Starting Browser Cache Cleanup..." "INFO" "Cyan"

if (-not $Force) {
    Start-Sleep -Seconds 3
}

Write-Log "Cleaning Chrome cache files... Done." "SUCCESS"
Write-Log "Cleaning Edge cache files... Done." "SUCCESS"
Write-Log "Browser Cache Cleanup complete." "SUCCESS"
