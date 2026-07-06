Param(
    [string]$Password = "B1g8lu3!"
)

# Standard UK format timestamped logger
$LogFolder = "C:\Seriun\log"
$LogFile   = "$LogFolder\VncPasswordReset.txt"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
    Write-Output "[ $TimeStamp ] $Message"
}

Write-Log "=== VNC Password Reset Initiated ==="

# Validate password length (VNC passwords are limited to 8 characters max due to DES key length limits)
if ($Password.Length -gt 8) {
    Write-Log "WARNING: VNC password exceeds 8 characters. Truncating to 8 characters."
    $Password = $Password.Substring(0, 8)
}

# 1. Encrypt password using standard VNC DES encryption
try {
    # Convert password to ASCII bytes and pad to 8 bytes with nulls
    $passBytes = [System.Text.Encoding]::ASCII.GetBytes($Password)
    $paddedBytes = [byte[]]::new(8)
    [array]::Copy($passBytes, $paddedBytes, $passBytes.Length)

    # TightVNC/VNC Magic DES Key (bit-reversed version of 17 52 6b 06 23 4e 58 07)
    $magicKey = [byte[]]@(0xE8, 0x4A, 0xD6, 0x60, 0xC4, 0x72, 0x1A, 0xE0)

    # Create DES encryptor
    $des = [System.Security.Cryptography.DES]::Create()
    $des.Padding = [System.Security.Cryptography.PaddingMode]::None
    $des.Mode = [System.Security.Cryptography.CipherMode]::ECB
    $des.Key = $magicKey

    $encryptor = $des.CreateEncryptor()
    $encryptedBytes = [byte[]]::new(8)
    $encryptor.TransformBlock($paddedBytes, 0, 8, $encryptedBytes, 0) | Out-Null
    
    $hexString = [System.BitConverter]::ToString($encryptedBytes).Replace("-", "")
    Write-Log "SUCCESS: Generated VNC encrypted password bytes (Hex: $hexString)"
} catch {
    Write-Log "ERROR: Encryption failed. Reason: $_"
    exit 1
}

# 2. Write to registry locations
$RegPaths = @(
    @{
        Path = "HKLM:\SOFTWARE\TightVNC\Server"
        Values = @{
            "Password" = $encryptedBytes
            "ControlPassword" = $encryptedBytes
            "UseControlAuthentication" = [int]1
            "UseViewerAuthentication" = [int]1
        }
    },
    @{
        Path = "HKLM:\SOFTWARE\ORL\WinVNC3"
        Values = @{
            "Password" = $encryptedBytes
        }
    },
    @{
        Path = "HKLM:\SOFTWARE\UltraVNC"
        Values = @{
            "Password" = $encryptedBytes
        }
    }
)

$registryUpdated = $false
foreach ($target in $RegPaths) {
    if (Test-Path $target.Path) {
        Write-Log "Found registry path: $($target.Path). Updating values..."
        try {
            foreach ($item in $target.Values.GetEnumerator()) {
                $valName = $item.Key
                $valData = $item.Value
                if ($valData -is [byte[]]) {
                    Set-ItemProperty -Path $target.Path -Name $valName -Value $valData -Type Binary -Force -ErrorAction Stop
                } elseif ($valData -is [int]) {
                    Set-ItemProperty -Path $target.Path -Name $valName -Value $valData -Type DWord -Force -ErrorAction Stop
                }
                Write-Log "  -> Set '$valName'"
            }
            $registryUpdated = $true
        } catch {
            Write-Log "  -> ERROR: Failed to write to $($target.Path). Reason: $_"
        }
    }
}

if (-not $registryUpdated) {
    Write-Log "WARNING: No standard VNC registry keys were found. VNC server might not be installed, or utilizes custom paths."
}

# 3. Restart VNC Services if running
$VncServices = Get-Service -Name "*tvn*", "*vnc*", "*tightvnc*", "*ultravnc*" -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "Running" }

if ($VncServices) {
    foreach ($service in $VncServices) {
        Write-Log "Attempting to restart VNC service: $($service.Name) ($($service.DisplayName))..."
        try {
            Restart-Service -Name $service.Name -Force -ErrorAction Stop
            Write-Log "SUCCESS: Service '$($service.Name)' restarted."
        } catch {
            Write-Log "ERROR: Failed to restart service '$($service.Name)'. Reason: $_"
        }
    }
} else {
    Write-Log "INFO: No running VNC services found to restart."
}

Write-Log "=== VNC Password Reset Finished ==="
exit 0
