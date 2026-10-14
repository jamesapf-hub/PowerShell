<#
.SYNOPSIS
    Downloads and silently installs the Datto RMM Agent (CagService) for a specified platform and site ID.
.DESCRIPTION
    Checks whether the Datto RMM Agent service ('CagService') is already present on the endpoint.
    If not installed, enforces TLS 1.2+, downloads the site-specific agent installer from the target
    Datto RMM platform (e.g. pinotage, merlot, concord, vidal, zinfandel, syrah), executes the silent
    installation, verifies the process exit code and service registration, and cleans up temporary files.
.PARAMETER Platform
    The Datto RMM platform subdomain (e.g., 'merlot', 'pinotage', 'concord', 'vidal', 'zinfandel', 'syrah').
    Defaults to $env:DATTO_PLATFORM or 'YOUR_DATTO_PLATFORM'.
.PARAMETER SiteID
    The Datto RMM Site UID / Profile ID for the target client site.
    Defaults to $env:DATTO_SITE_ID or 'YOUR_DATTO_SITE_ID'.
.EXAMPLE
    .\Install-DattoRmmAgent.ps1 -Platform "merlot" -SiteID "12345678-abcd-1234-efab-1234567890ab"
.NOTES
    Author     : James
    Version    : 2.0
    Log Path   : $env:SystemDrive\Logs\Install-DattoRmmAgent\DattoAgentInstall.txt
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Platform = $(if ($env:DATTO_PLATFORM) { $env:DATTO_PLATFORM } else { "YOUR_DATTO_PLATFORM" }),

    [Parameter(Mandatory = $false)]
    [string]$SiteID = $(if ($env:DATTO_SITE_ID) { $env:DATTO_SITE_ID } else { "YOUR_DATTO_SITE_ID" })
)

$LogFolder = "$env:SystemDrive\Logs\Install-DattoRmmAgent"
$LogFile   = "$LogFolder\DattoAgentInstall.txt"

if (-not (Test-Path -Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $Entry = "[ $TimeStamp ] $Message"
    $Entry | Out-File -FilePath $LogFile -Append -Encoding utf8
    Write-Output $Entry
}

Write-Log "=== Datto RMM Agent Deployment Started ==="

# 1. Verify if Datto RMM Agent is already installed
$existingService = Get-Service -Name "CagService" -ErrorAction SilentlyContinue
if ($null -ne $existingService) {
    Write-Log "INFO: Datto RMM Agent (CagService) is already installed (Status: $($existingService.Status)). Exiting."
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Platform) -or $Platform -eq "YOUR_DATTO_PLATFORM" -or
    [string]::IsNullOrWhiteSpace($SiteID) -or $SiteID -eq "YOUR_DATTO_SITE_ID") {
    Write-Log "ERROR: Platform or SiteID has not been configured. Supply -Platform and -SiteID or set DATTO_PLATFORM / DATTO_SITE_ID environment variables."
    exit 1
}

# 2. Enforce TLS 1.2+ for secure download
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {
    $AvailableProtocols = [enum]::GetNames([Net.SecurityProtocolType]) -join ", "
    Write-Log "ERROR: Unable to configure TLS 1.2 security protocol. Available protocols: $AvailableProtocols."
    exit 1
}

# 3. Download the Agent installer
$AgentURL = "https://$Platform.centrastage.net/csm/profile/downloadAgent/$SiteID"
$InstallerPath = Join-Path -Path $env:TEMP -ChildPath "DRMMSetup.exe"
$DownloadStart = Get-Date

Write-Log "INFO: Starting Agent download from $AgentURL"

try {
    Invoke-WebRequest -Uri $AgentURL -OutFile $InstallerPath -UseBasicParsing -ErrorAction Stop
    if (-not (Test-Path -Path $InstallerPath) -or (Get-Item -Path $InstallerPath).Length -lt 10240) {
        throw "Downloaded installer file is missing or unexpectedly small."
    }
} catch {
    Write-Log "ERROR: Agent installer download failed. Reason: $($_.Exception.Message)"
    exit 1
}

$DownloadSeconds = [Math]::Round(((Get-Date) - $DownloadStart).TotalSeconds, 1)
Write-Log "SUCCESS: Agent download completed in $DownloadSeconds seconds."

# 4. Install the Agent silently and validate exit code
$InstallStart = Get-Date
Write-Log "INFO: Executing Datto RMM Agent installer..."

try {
    $Process = Start-Process -FilePath $InstallerPath -PassThru -Wait -WindowStyle Hidden -ErrorAction Stop
    $InstallSeconds = [Math]::Round(((Get-Date) - $InstallStart).TotalSeconds, 1)

    if ($Process.ExitCode -ne 0) {
        Write-Log "ERROR: Installer exited with non-zero code $($Process.ExitCode) after $InstallSeconds seconds."
        exit $Process.ExitCode
    }

    Write-Log "SUCCESS: Agent installation process finished in $InstallSeconds seconds (ExitCode: 0)."
} catch {
    Write-Log "ERROR: Agent installation failed to execute. Reason: $($_.Exception.Message)"
    exit 1
} finally {
    if (Test-Path -Path $InstallerPath) {
        Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        Write-Log "INFO: Cleaned up temporary installer file ($InstallerPath)."
    }
}

# 5. Verify service registration post-install
Start-Sleep -Seconds 3
$VerifyService = Get-Service -Name "CagService" -ErrorAction SilentlyContinue
if ($null -ne $VerifyService) {
    Write-Log "SUCCESS: Verified 'CagService' is registered and $($VerifyService.Status)."
    Write-Log "=== Datto RMM Agent Deployment Finished ==="
    exit 0
} else {
    Write-Log "WARNING: Installer completed, but 'CagService' was not immediately detected."
    Write-Log "=== Datto RMM Agent Deployment Finished (With Warning) ==="
    exit 0
}