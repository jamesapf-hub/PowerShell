$Platform="YOUR_DATTO_PLATFORM"
$SiteID="YOUR_DATTO_SITE_ID" 

# Standard UK format timestamped logger
$LogFolder = "$env:SystemDrive\Logs\Install-DattoRmmAgent"
$LogFile   = "$LogFolder\DattoAgentInstall.txt"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
    Write-Output "[ $TimeStamp ] $Message"
}

Write-Log "=== Datto RMM Agent Deployment Started ==="

# First check if Agent is installed and instantly exit if so
if (Get-Service CagService -ErrorAction SilentlyContinue) {
    Write-Log "INFO: Datto RMM Agent already installed on this device. Exiting."
    exit 0
} 

# Download the Agent
$AgentURL="https://$Platform.centrastage.net/csm/profile/downloadAgent/$SiteID" 
$DownloadStart=Get-Date 
Write-Log "Starting Agent download from $AgentURL"

try {
    [Net.ServicePointManager]::SecurityProtocol=[Enum]::ToObject([Net.SecurityProtocolType],3072)
}
catch {
    $AvailableProtocols = [enum]::GetNames([Net.SecurityProtocolType]) -join ", "
    Write-Log "ERROR: Cannot download Agent due to invalid security protocol. Available protocols: $AvailableProtocols. Agent download requires at least TLS 1.2 to succeed. Please install TLS 1.2 and rerun the script."
    exit 1
}

try {
    (New-Object System.Net.WebClient).DownloadFile($AgentURL, "$env:TEMP\DRMMSetup.exe")
} 
catch {
    Write-Log "ERROR: Agent installer download failed. Exit message: $_"
    exit 1
} 

$DownloadTime = ((Get-Date).Subtract($DownloadStart).TotalSeconds).ToString("F0")
Write-Log "SUCCESS: Agent download completed in $DownloadTime seconds." 

# Install the Agent
$InstallStart=Get-Date 
Write-Log "Starting Agent installation..." 
try {
    & "$env:TEMP\DRMMSetup.exe" | Out-Null 
    $InstallTime = ((Get-Date).Subtract($InstallStart).TotalSeconds).ToString("F0")
    Write-Log "SUCCESS: Agent install completed in $InstallTime seconds."
} catch {
    Write-Log "ERROR: Agent installation failed. Reason: $_"
}

# Cleanup installer
if (Test-Path "$env:TEMP\DRMMSetup.exe") {
    Remove-Item "$env:TEMP\DRMMSetup.exe" -Force | Out-Null
    Write-Log "Cleanup: Removed setup file from temp."
}

Write-Log "=== Datto RMM Agent Deployment Finished ==="
exit 0