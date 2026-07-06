# Define the permanent redirect URL for the latest Edge Stable x64 MSI
$edgeInstallerUrl = "https://go.microsoft.com/fwlink/?LinkID=2093437"

# Set the local path for the installer
$installerPath = "$env:TEMP\MicrosoftEdgeEnterpriseX64.msi"

# Standard UK format timestamped logger
$LogFolder = "$env:SystemDrive\Logs\Update-EdgeBrowser"
$LogFile   = "$LogFolder\EdgeUpdate.txt"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Msg, $Color = "White") { 
    $TimeStamp = Get-Date -Format 'dd/MM/yy HH:mm:ss'
    "[ $TimeStamp ] $Msg" | Out-File $LogFile -Append 
    Write-Host "[*] $Msg" -ForegroundColor $Color
}

Write-Log "=== Microsoft Edge Installation/Update Started ==="

# Validate Administrator Elevation
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "ERROR: This script must be run as an Administrator (Elevated PowerShell session)." "Red"
    exit 1
}

# Enable TLS 1.2 for the download session
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Log "WARNING: Failed to force TLS 1.2. The download might fail if the server requires it. Reason: $_" "Yellow"
}

Write-Log "Downloading Microsoft Edge installer from $edgeInstallerUrl ..." "Cyan"

try {
    # Use -UseBasicParsing to prevent Internet Explorer first-run initialization errors
    Invoke-WebRequest -Uri $edgeInstallerUrl -OutFile $installerPath -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Log "ERROR: Failed to download the Edge installer. Reason: $_" "Red"
    exit 1
}

# Check if the download succeeded
if (Test-Path $installerPath) {
    Write-Log "Download complete. Installing/Updating Edge..." "Cyan"

    try {
        # Run the installer silently and capture the process to check exit code
        $Process = Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait -PassThru -ErrorAction Stop
        $ExitCode = $Process.ExitCode

        # Exit code 0 = Success, 3010 = Success with reboot required, 1641 = Success with reboot initiated
        if ($ExitCode -eq 0 -or $ExitCode -eq 3010 -or $ExitCode -eq 1641) {
            Write-Log "SUCCESS: Microsoft Edge has been updated. Exit code: $ExitCode" "Green"
        } else {
            Write-Log "ERROR: msiexec failed with exit code: $ExitCode" "Red"
        }
    } catch {
        Write-Log "ERROR: Failed to launch msiexec. Reason: $_" "Red"
    } finally {
        # Clean up installer
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force | Out-Null
            Write-Log "Cleanup: Removed installer file from temp."
        }
    }
} else {
    Write-Log "ERROR: Installer file not found in Temp directory." "Red"
}

Write-Log "=== Microsoft Edge Installation/Update Finished ==="
exit 0
