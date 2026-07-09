<#
.SYNOPSIS
Securely enables a Microsoft Teams Room (MTR) mailbox to accept forwarded external meeting invites while blocking direct bookings from outside the organization.

.DESCRIPTION
This script applies two crucial settings to the specified room mailbox:
1. Blocks all unauthenticated external senders from booking or emailing the room (security).
2. Enables processing of external meeting content and link retention for internal users (functionality).
It also logs all actions to $env:SystemDrive\Logs\ExternalTeamsRoom\ExternalTeamsRoom.log.

.PARAMETER RoomEmail
The email address of the room mailbox to configure.

.NOTES
Author      : JP
Created     : 2025-10-22
Version     : 1.1
#>
param (
    [Parameter(Mandatory = $true)]
    [string] $RoomEmail = "room@contoso.com"
)

# --- CONFIGURATION & LOGGING ---
$LogDir = "$env:SystemDrive\Logs\ExternalTeamsRoom"
$LogFile = "$LogDir\ExternalTeamsRoom.log"

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
    
    # Write to local log file
    Add-Content -Path $LogFile -Value $LogEntry
    
    # Write to console with colors
    switch ($Level) {
        "INFO"    { Write-Host $LogEntry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        default   { Write-Host $LogEntry }
    }
}

Write-Log "Starting ExternalTeamsRoom configuration for: $RoomEmail" "INFO"

# --- 1. CONNECT TO EXCHANGE ONLINE ---
Write-Log "Checking for ExchangeOnlineManagement module..." "INFO"
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Log "ExchangeOnlineManagement module not found." "WARNING"
    $installPrompt = Read-Host "Would you like to install the ExchangeOnlineManagement module now? (Y/N)"
    if ($installPrompt -match '^[Yy]$') {
        Write-Log "Installing ExchangeOnlineManagement module. This may take a few moments..." "INFO"
        try {
            # Requires NuGet provider. Force suppresses prompts.
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -ErrorAction Stop
            Write-Log "Module installed successfully." "SUCCESS"
        } catch {
            Write-Log "Failed to install the module. Please run PowerShell as Administrator and try again, or install it manually." "ERROR"
            Write-Log $_.Exception.Message "ERROR"
            exit
        }
    } else {
        Write-Log "Module installation declined. Exiting script." "ERROR"
        exit
    }
}

Write-Log "Connecting to Exchange Online..." "INFO"
try {
    # Use Connect-ExchangeOnline with your preferred authentication method
    Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
    Write-Log "Successfully connected to Exchange Online." "SUCCESS"
} catch {
    Write-Log "Failed to connect to Exchange Online. Ensure you have sufficient permissions." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
    exit
}

# --- 2. ENFORCE SECURITY (Block Direct External Mail) ---
Write-Log "Setting -RequireSenderAuthenticationEnabled to '$True' for $RoomEmail to block direct external mail." "INFO"
try {
    Set-Mailbox -Identity $RoomEmail -RequireSenderAuthenticationEnabled $True -ErrorAction Stop
    Write-Log "Security restriction applied successfully." "SUCCESS"
} catch {
    Write-Log "Failed to set mailbox security restrictions." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
}

# --- 3. ENABLE FUNCTIONALITY (Allow Forwarded Invites) ---
Write-Log "Setting Calendar Processing properties to allow forwarded external invites and preserve meeting details." "INFO"
try {
    Set-CalendarProcessing -Identity $RoomEmail -ProcessExternalMeetingMessages $True -DeleteComments $False -ErrorAction Stop
    Write-Log "Calendar processing successfully configured." "SUCCESS"
} catch {
    Write-Log "Failed to set calendar processing properties." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
}

# --- 4. DISCONNECT ---
Write-Log "Configuration complete. Disconnecting from Exchange Online." "INFO"
Disconnect-ExchangeOnline -Confirm:$false
Write-Log "Disconnected successfully." "SUCCESS"
