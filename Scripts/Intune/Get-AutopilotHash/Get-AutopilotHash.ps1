<#
.SYNOPSIS
    Harvests the Windows Autopilot Hardware Hash (HWID) and uploads it to phnx.it for instant short-URL CSV download on another machine.

.DESCRIPTION
    Queries the local Windows Management Instrumentation (WMI/CIM) MDM_DevDetail_Ext01 class natively (zero PSGallery/NuGet dependencies,
    ideal for Shift+F10 OOBE provisioning) to extract the Device Serial Number and 4KB Hardware Hash.
    Saves a local CSV backup to C:\HWID, logs execution details to C:\Logs\Get-AutopilotHash, and uploads the payload to the
    phnx.it Cloudflare Edge endpoint. Outputs a short URL (e.g. https://phnx.it/ap/K9X4M or https://phnx.it/BENCH1) that immediately
    downloads an Intune-ready Autopilot CSV file on any admin workstation, supporting multi-device Batch PIN grouping.

.PARAMETER BatchPin
    Optional alphanumeric session/batch code (e.g. 'BENCH1' or '1234'). Any devices uploaded with the same BatchPin within 24 hours
    are automatically combined and deduplicated into a single downloadable Intune CSV file.

.PARAMETER GroupTag
    Optional Microsoft Intune Autopilot Group Tag to append as the 4th CSV column ('Group Tag').

.PARAMETER SkipPrompt
    Bypasses interactive console prompts for BatchPin and GroupTag (useful for zero-touch automated execution).

.EXAMPLE
    irm phnx.it/APHASH | iex

.EXAMPLE
    .\Get-AutopilotHash.ps1 -BatchPin "BENCH1" -GroupTag "StandardUser" -SkipPrompt

.NOTES
    Author:  Seriun IT / Helper Portal
    Version: 1.0.0
    Logs:    C:\Logs\Get-AutopilotHash\AutopilotHash_<ddMMyy-HHmmss>.log
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$BatchPin = $env:AP_PIN,

    [Parameter(Mandatory = $false)]
    [string]$GroupTag = $env:AP_TAG,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPrompt
)

# Enforce TLS 1.2 for Windows PowerShell 5.1 / OOBE environments
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Initialize standardized logging directory with local AppData fallback
$logDir = "$env:SystemDrive\Logs\Get-AutopilotHash"
try {
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    $logDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Get-AutopilotHash"
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

$timestamp = Get-Date -Format "ddMMyy-HHmmss"
$logPath = Join-Path -Path $logDir -ChildPath "AutopilotHash_$timestamp.log"

function Write-AutopilotLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timeStr] [$Level] $Message"
    try {
        $entry | Out-File -FilePath $logPath -Append -Encoding utf8 -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore file write errors in read-only PE environments
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   WINDOWS AUTOPILOT HARDWARE HASH HARVESTER (phnx.it)      " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-AutopilotLog -Message "Started Windows Autopilot Hardware Hash extraction."

# Verify Administrator privileges (required to read root/cimv2/mdm/dmmap)
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "[!] WARNING: Not running as Administrator." -ForegroundColor Yellow
    Write-Host "    Querying MDM_DevDetail_Ext01 requires an elevated PowerShell prompt (or Shift+F10 in OOBE)." -ForegroundColor Yellow
    Write-AutopilotLog -Message "Script launched without Administrator privileges." -Level "WARN"
}

try {
    # 1. Interactive Prompts for Optional Batch PIN & Group Tag
    if (-not $SkipPrompt -and [Environment]::UserInteractive) {
        if ([string]::IsNullOrWhiteSpace($BatchPin)) {
            Write-Host ""
            Write-Host "Optional: Enter a Batch PIN (e.g. BENCH1) to combine multiple PCs into one CSV." -ForegroundColor Gray
            $inputPin = Read-Host "Batch PIN (Press Enter for single PC link)"
            if (-not [string]::IsNullOrWhiteSpace($inputPin)) {
                $BatchPin = $inputPin.Trim()
            }
        }
        if ([string]::IsNullOrWhiteSpace($GroupTag)) {
            $inputTag = Read-Host "Intune Group Tag (Press Enter to skip)"
            if (-not [string]::IsNullOrWhiteSpace($inputTag)) {
                $GroupTag = $inputTag.Trim()
            }
        }
    }

    Write-Host ""
    Write-Host "[*] Querying BIOS and MDM Hardware Hash via CIM..." -ForegroundColor Cyan

    # 2. Query System & BIOS Information
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue

    $serialNumber = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "UNKNOWN" }
    $hostname = if ($cs.Name) { $cs.Name.Trim() } else { $env:COMPUTERNAME }
    $manufacturer = if ($cs.Manufacturer) { $cs.Manufacturer.Trim() } else { "" }
    $model = if ($cs.Model) { $cs.Model.Trim() } else { "" }

    Write-AutopilotLog -Message "Detected Hostname: $hostname | Serial: $serialNumber | Model: $manufacturer $model"

    # 3. Query MDM_DevDetail_Ext01 for the 4KB Autopilot Hardware Hash
    $devDetail = Get-CimInstance -Namespace "root/cimv2/mdm/dmmap" -ClassName "MDM_DevDetail_Ext01" -Filter "InstanceID='Ext' AND ParentID='./DevDetail'" -ErrorAction Stop

    if (-not $devDetail -or [string]::IsNullOrWhiteSpace($devDetail.DeviceHardwareData)) {
        throw "MDM_DevDetail_Ext01 returned an empty hardware hash. Ensure you are running PowerShell as Administrator on Windows 10/11."
    }

    $hardwareHash = $devDetail.DeviceHardwareData.Trim()
    Write-AutopilotLog -Message "Successfully extracted Hardware Hash (Length: $($hardwareHash.Length) chars)." -Level "SUCCESS"

    # 4. Save Local Backup CSV in C:\HWID
    $hwidDir = "$env:SystemDrive\HWID"
    try {
        if (-not (Test-Path -Path $hwidDir)) {
            New-Item -ItemType Directory -Path $hwidDir -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        $hwidDir = $logDir
    }

    $localCsvPath = Join-Path -Path $hwidDir -ChildPath "AutopilotHWID-$serialNumber.csv"
    if (-not [string]::IsNullOrWhiteSpace($GroupTag)) {
        $csvHeader = "Device Serial Number,Windows Product ID,Hardware Hash,Group Tag"
        $csvRow = "$serialNumber,,$hardwareHash,$GroupTag"
    }
    else {
        $csvHeader = "Device Serial Number,Windows Product ID,Hardware Hash"
        $csvRow = "$serialNumber,,$hardwareHash"
    }

    @($csvHeader, $csvRow) | Out-File -FilePath $localCsvPath -Encoding ascii -Force
    Write-AutopilotLog -Message "Saved local Autopilot CSV backup to: $localCsvPath"

    # 5. Upload Hardware Hash to phnx.it Cloud Endpoint
    Write-Host "[*] Uploading Hardware Hash to phnx.it cloud drop..." -ForegroundColor Cyan

    $payloadObj = @{
        serialNumber = $serialNumber
        productId    = ""
        hardwareHash = $hardwareHash
        groupTag     = $GroupTag
        hostname     = $hostname
        manufacturer = $manufacturer
        model        = $model
        batchPin     = $BatchPin
    }
    $jsonPayload = $payloadObj | ConvertTo-Json -Depth 3 -Compress

    $endpoints = @(
        "https://phnx.it/api/autopilot/upload",
        "https://helper.pheneox.net/api/autopilot/upload"
    )

    $response = $null
    $lastErr = $null

    foreach ($uri in $endpoints) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Post -Body $jsonPayload -ContentType "application/json; charset=utf-8" -TimeoutSec 15 -ErrorAction Stop
            if ($response -and $response.success) {
                break
            }
        }
        catch {
            $lastErr = $_
            Write-AutopilotLog -Message "Upload attempt to $uri failed: $($_.Exception.Message)" -Level "WARN"
        }
    }

    if (-not $response -or -not $response.success) {
        throw "Cloud upload failed ($($lastErr.Exception.Message)). Local CSV is still saved at: $localCsvPath"
    }

    $code = $response.code
    $downloadUrl = $response.downloadUrl
    $shortUrl = $response.shortDownloadUrl
    $deviceCount = $response.deviceCount

    Write-AutopilotLog -Message "Cloud upload succeeded! Code: $code | Download URL: $downloadUrl | Devices in batch: $deviceCount" -Level "SUCCESS"

    # 6. Output Summary & Short Download Links
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " [+] AUTOPILOT HARDWARE HASH READY FOR DOWNLOAD!" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host (" {0,-18} : {1}" -f "Hostname", $hostname) -ForegroundColor White
    Write-Host (" {0,-18} : {1}" -f "Model", "$manufacturer $model".Trim()) -ForegroundColor White
    Write-Host (" {0,-18} : {1}" -f "Serial Number", $serialNumber) -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($GroupTag)) {
        Write-Host (" {0,-18} : {1}" -f "Group Tag", $GroupTag) -ForegroundColor White
    }
    Write-Host (" {0,-18} : {1} device(s) in CSV (Session: {2})" -f "Batch Status", $deviceCount, $code) -ForegroundColor Cyan
    Write-Host (" {0,-18} : {1}" -f "Local Backup", $localCsvPath) -ForegroundColor DarkGray
    Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host " Open either link on your other machine to download the CSV:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   $downloadUrl" -ForegroundColor Cyan
    Write-Host "   $shortUrl" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " (Link valid for 24 hours. Run with Batch PIN '$code' on more PCs to combine!)" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
}
catch {
    $errMsg = $_.Exception.Message
    Write-AutopilotLog -Message "Fatal error: $errMsg" -Level "ERROR"
    Write-Host ""
    Write-Host "[X] Failed to harvest or upload Autopilot Hardware Hash:" -ForegroundColor Red
    Write-Host "    $errMsg" -ForegroundColor Red
    Write-Host "    Log saved to: $logPath" -ForegroundColor DarkGray
    Write-Host ""
}
