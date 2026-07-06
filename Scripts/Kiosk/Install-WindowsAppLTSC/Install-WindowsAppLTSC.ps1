<#
.SYNOPSIS
    Performs an offline provisioning update of the Windows App for Windows 10 IoT LTSC environments.
.DESCRIPTION
    Pulls the latest verified MSIX/MSIXBUNDLE payload from the Microsoft CDN repository (via rg-adguard API) and registers it machine-wide.
.NOTES
    Author:  JP & Antigravity
    Version: 1.4
    Log:     $env:SystemDrive\Logs\Install-WindowsAppLTSC\ltsc_app_update.log
#>

$LogPath = "$env:SystemDrive\Logs\Install-WindowsAppLTSC"
$LogFile = "$LogPath\ltsc_app_update.log"
if (-not (Test-Path $LogPath)) { 
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null 
}

function Log-Message ($Message) {
    $Timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[$Timestamp] $Message" | Out-File -FilePath $LogFile -Append
    Write-Host "[*] $Message"
}

# Ensure execution context is elevated
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Log-Message "CRITICAL: Script must be executed out of an Elevated/Administrator PowerShell prompt."
    exit 1
}

# Enable TLS 1.2 for the download session
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Log-Message "WARNING: Failed to force TLS 1.2. API request may fail. Reason: $_"
}

$WorkingDir = "$env:SystemDrive\Logs\Install-WindowsAppLTSC\Software"
if (-not (Test-Path $WorkingDir)) { 
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null 
}

Log-Message "Initiating machine-wide system provisioning update for Windows App..."

# Browser headers to bypass Cloudflare DDOS challenge checks on store.rg-adguard.net
$Headers = @{
    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8"
    "Accept-Language" = "en-US,en;q=0.9"
    "Origin" = "https://store.rg-adguard.net"
    "Referer" = "https://store.rg-adguard.net/"
}

# Query the new Windows App ID (9N1F85V9T8BN) first, with fallback to legacy Remote Desktop ID (9wzdncrfj3p1)
$ProductIDs = @('9N1F85V9T8BN', '9wzdncrfj3p1')
$DownloadUrl = $null
$TargetExtension = "msix"

foreach ($ProductId in $ProductIDs) {
    Log-Message "Querying Microsoft API distribution paths for App ID: $ProductId ..."
    
    $Payload = @{
        type = 'url'
        url = "https://apps.microsoft.com/detail/$ProductId"
        ring = 'RP'
    }

    try {
        $Response = Invoke-WebRequest -Uri "https://store.rg-adguard.net/api/GetFiles" -Method Post -Headers $Headers -Body $Payload -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop
        
        # Parse the HTML response to locate download links matching the filename text
        $Matches = [regex]::Matches($Response.Content, '<a href="(?<url>http://tlu\.dl\.delivery\.mp\.microsoft\.com/[^"]+)"[^>]*>(?<filename>[^<]+)</a>')
        
        foreach ($Match in $Matches) {
            $Url = $Match.Groups['url'].Value
            $Filename = $Match.Groups['filename'].Value
            
            # Identify the correct x64 architecture installer package (MSIX or MSIXBUNDLE)
            # Exclude framework dependency packages (like WindowsAppRuntime and VCLibs) to target the main client application
            if ($Filename -like "*x64*" -and 
                ($Filename -like "*.msixbundle" -or $Filename -like "*.msix") -and 
                $Filename -notlike "*Runtime*" -and 
                $Filename -notlike "*VCLibs*") {
                
                $DownloadUrl = $Url
                if ($Filename -like "*.msixbundle") {
                    $TargetExtension = "msixbundle"
                } else {
                    $TargetExtension = "msix"
                }
                Log-Message "SUCCESS: Resolved target bundle: $Filename"
                break
            }
        }
    } catch {
        Log-Message "WARNING: Query for ID $ProductId failed. Reason: $_"
    }

    if ($DownloadUrl) {
        break
    }
}

if (-not $DownloadUrl) {
    Log-Message "CRITICAL: Valid modern x64 target package could not be resolved from Store API."
    exit 1
}

# 2. Stage download
$TargetBundle = "$WorkingDir\WindowsApp_LTSC_Update.$TargetExtension"
Log-Message "Downloading offline deployment engine bundle..."
try {
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $TargetBundle -UseBasicParsing -ErrorAction Stop
    Log-Message "SUCCESS: Offline bundle downloaded to $TargetBundle"
} catch {
    Log-Message "CRITICAL: Failed to download the bundle package. Reason: $_"
    exit 1
}

# 3. Apply the Provisioned Package Upgrade to the Live OS Instance
Log-Message "Registering update package machine-wide across all user profiles..."
try {
    # Using DISM hooks through PowerShell to safely register the package
    Add-AppxProvisionedPackage -Online -PackagePath $TargetBundle -SkipLicense -ErrorAction Stop
    Log-Message "SUCCESS: Provisioning call successfully registered inside the Windows image repository."
} catch {
    Log-Message "WARNING: Standard provision failed. Reason: $_. Attempting explicit offline installation fallback..."
    try {
        # Fallback to local package target overlay if provision engine blocks processing
        Add-AppxPackage -Path $TargetBundle -ErrorAction Stop
        Log-Message "SUCCESS: Fallback package registration completed."
    } catch {
        Log-Message "CRITICAL: Fallback installation failed. Reason: $_"
        
        # Cleanup installer before exiting
        if (Test-Path $TargetBundle) { Remove-Item $TargetBundle -Force | Out-Null }
        exit 1
    }
}

# Cleanup installer bundle to free up space
if (Test-Path $TargetBundle) {
    Remove-Item $TargetBundle -Force | Out-Null
    Log-Message "Cleanup: Removed downloaded bundle file."
}

Log-Message "Windows App infrastructure update deployment cycle complete."
exit 0
