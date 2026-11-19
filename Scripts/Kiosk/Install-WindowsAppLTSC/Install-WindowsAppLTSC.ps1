<#
.SYNOPSIS
    Performs an offline provisioning update of the Windows App for Windows 10/11 IoT LTSC environments.
.DESCRIPTION
    Downloads the official Microsoft Windows App x64 MSIX/MSI installer package directly from Microsoft's official distribution endpoint, verifies its Authenticode digital signature, and provisions it machine-wide.
.PARAMETER PackageUrl
    Official Microsoft direct distribution URL for the Windows App x64 package.
.NOTES
    Author:  JP
    Version: 1.5
    Log:     $env:SystemDrive\Logs\Install-WindowsAppLTSC\ltsc_app_update.log
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$PackageUrl = "https://go.microsoft.com/fwlink/?linkid=2262633"
)

$LogPath = "$env:SystemDrive\Logs\Install-WindowsAppLTSC"
$LogFile = "$LogPath\ltsc_app_update.log"
if (-not (Test-Path $LogPath)) { 
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null 
}

function Write-LTSCLog ($Message) {
    $Timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[$Timestamp] $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
    Write-Host "[*] $Message"
}

# Ensure execution context is elevated
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-LTSCLog "CRITICAL: Script must be executed out of an Elevated/Administrator PowerShell prompt."
    exit 1
}

# Enable TLS 1.2 for the download session
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-LTSCLog "WARNING: Failed to force TLS 1.2. Request may fail. Reason: $_"
}

$WorkingDir = "$env:SystemDrive\Logs\Install-WindowsAppLTSC\Software"
if (-not (Test-Path $WorkingDir)) { 
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null 
}

Write-LTSCLog "Initiating machine-wide system provisioning update for Microsoft Windows App..."

# 1. Stage download from official Microsoft endpoint
$TargetBundle = "$WorkingDir\WindowsApp_LTSC_Update.msix"
Write-LTSCLog "Downloading official Windows App bundle from $PackageUrl ..."
try {
    Invoke-WebRequest -Uri $PackageUrl -OutFile $TargetBundle -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
    Write-LTSCLog "SUCCESS: Offline bundle downloaded to $TargetBundle"
} catch {
    Write-LTSCLog "CRITICAL: Failed to download the bundle package. Reason: $_"
    exit 1
}

# 2. Verify Authenticode digital signature before provisioning
try {
    $Sig = Get-AuthenticodeSignature -FilePath $TargetBundle -ErrorAction Stop
    if ($Sig.Status -ne 'Valid') {
        throw "Authenticode signature status is '$($Sig.Status)'."
    }
    Write-LTSCLog "SUCCESS: Verified Microsoft Authenticode signature ($($Sig.SignerCertificate.Subject))."
} catch {
    Write-LTSCLog "CRITICAL: Authenticode verification failed for $TargetBundle. Reason: $_"
    if (Test-Path $TargetBundle) {
        try { Remove-Item $TargetBundle -Force -ErrorAction Stop } catch { Write-LTSCLog "WARNING: Cleanup failed: $_" }
    }
    exit 1
}

# 3. Apply the Provisioned Package Upgrade to the Live OS Instance
Write-LTSCLog "Registering update package machine-wide across all user profiles..."
try {
    Add-AppxProvisionedPackage -Online -PackagePath $TargetBundle -SkipLicense -ErrorAction Stop
    Write-LTSCLog "SUCCESS: Provisioning call successfully registered inside the Windows image repository."
} catch {
    Write-LTSCLog "WARNING: Standard provision failed. Reason: $_. Attempting explicit offline installation fallback..."
    try {
        Add-AppxPackage -Path $TargetBundle -ErrorAction Stop
        Write-LTSCLog "SUCCESS: Fallback package registration completed."
    } catch {
        Write-LTSCLog "CRITICAL: Fallback installation failed. Reason: $_"
        if (Test-Path $TargetBundle) {
            try { Remove-Item $TargetBundle -Force -ErrorAction Stop } catch { Write-LTSCLog "WARNING: Cleanup failed: $_" }
        }
        exit 1
    }
}

# Cleanup installer bundle to free up space
if (Test-Path $TargetBundle) {
    try {
        Remove-Item $TargetBundle -Force -ErrorAction Stop
        Write-LTSCLog "Cleanup: Removed downloaded bundle file."
    } catch {
        Write-LTSCLog "WARNING: Failed to remove temporary bundle file: $_"
    }
}

Write-LTSCLog "Windows App infrastructure update deployment cycle complete."
exit 0
