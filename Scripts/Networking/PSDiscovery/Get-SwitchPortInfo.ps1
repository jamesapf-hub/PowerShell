<#
.SYNOPSIS
    Queries the local network adapter for CDP/LLDP packets to identify the connected switch port.
.DESCRIPTION
    This script checks for and installs the PSDiscoveryProtocol module if missing,
    captures discovery packets, and outputs the switch name, port, VLAN, and model.
    It logs progress to C:\Log\PSDiscovery\ and appends network info to a central CSV file.
.EXAMPLE
    .\Get-SwitchPortInfo.ps1
#>

# 1. Locate and import the Logging module dynamically
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CandidatePaths = @(
    (Join-Path $ScriptDir "..\Modules\Logging\Logging.psm1"),
    (Join-Path $ScriptDir "Modules\Logging\Logging.psm1"),
    (Join-Path $ScriptDir "..\..\Modules\Logging\Logging.psm1"),
    (Join-Path $ScriptDir "..\..\..\Modules\Logging\Logging.psm1") # Safe fallback for 3 levels nesting
)

$ModuleImported = $false
foreach ($Path in $CandidatePaths) {
    if (Test-Path $Path) {
        Import-Module $Path -Force -ErrorAction SilentlyContinue
        $ModuleImported = $true
        break
    }
}

if (-not $ModuleImported) {
    Write-Error "CRITICAL: The Logging module (Logging.psm1) could not be located."
    exit 1
}

# Define the log and inventory output directory
$OutputDir = "C:\Log\PSDiscovery"
$CsvPath = Join-Path $OutputDir "SwitchInventory.csv"

# 2. Initialize logging (creates C:\Log\PSDiscovery if it doesn't exist)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = $MyInvocation.MyCommand.Path
$LogFile = Initialize-ScriptLogging -ScriptPath $ScriptPath -LogDirectory $OutputDir

try {
    Write-ScriptLog -Message "Starting Get-SwitchPortInfo execution." -LogFile $LogFile -Level INFO

    # 3. Check and install PSDiscoveryProtocol dependency
    if (-not (Get-Module -ListAvailable -Name PSDiscoveryProtocol)) {
        Write-ScriptLog -Message "PSDiscoveryProtocol module is not installed. Attempting installation for CurrentUser..." -LogFile $LogFile -Level WARNING
        try {
            Install-Module -Name PSDiscoveryProtocol -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-ScriptLog -Message "PSDiscoveryProtocol installed successfully." -LogFile $LogFile -Level SUCCESS
        } 
        catch {
            Write-ScriptLog -Message "Auto-installation failed. Please run 'Install-Module PSDiscoveryProtocol -Scope CurrentUser' in an elevated console." -LogFile $LogFile -Level ERROR
            throw $_
        }
    }

    # 4. Perform Packet Capture
    Write-ScriptLog -Message "Listening for CDP/LLDP packets on network interfaces (this may take up to 60 seconds)..." -LogFile $LogFile -Level INFO

    $Packet = Invoke-DiscoveryProtocolCapture -ErrorAction Stop

    if ($Packet) {
        Write-ScriptLog -Message "Packet(s) captured successfully. Parsing discovery data..." -LogFile $LogFile -Level INFO
        $Data = Get-DiscoveryProtocolData -Packet $Packet

        if ($Data) {
            $Records = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($Connection in $Data) {
                $Device = $Connection.Device
                $Port = $Connection.Port
                $VLAN = $Connection.VLAN
                $Model = $Connection.Model
                $IP = $Connection.IPAddress
                $Type = $Connection.Type

                # Format human-readable console/log message
                $InfoMsg = "[$Type] Connected to Device: $Device | Port: $Port | VLAN: $VLAN | Switch IP: $IP | Switch Model: $Model"
                Write-ScriptLog -Message $InfoMsg -LogFile $LogFile -Level SUCCESS

                # Build record object for CSV export
                $Record = [PSCustomObject]@{
                    Timestamp    = (Get-Date -Format "dd/MM/yy HH:mm:ss")
                    ComputerName = $env:COMPUTERNAME
                    ProtocolType = $Type
                    SwitchDevice = $Device
                    SwitchPort   = $Port
                    VLAN         = $VLAN
                    SwitchIP     = $IP
                    SwitchModel  = $Model
                }
                $Records.Add($Record)
            }

            # 5. Export/Append findings to the CSV file
            try {
                $Records | Export-Csv -Path $CsvPath -Append -NoTypeInformation -ErrorAction Stop
                Write-ScriptLog -Message "Appended discovery findings to CSV inventory: $CsvPath" -LogFile $LogFile -Level SUCCESS
            }
            catch {
                Write-ScriptLog -Message "Failed to write results to CSV file. Error: $_" -LogFile $LogFile -Level ERROR
            }
        }
        else {
            Write-ScriptLog -Message "Captured packets, but no readable CDP or LLDP discovery data could be extracted." -LogFile $LogFile -Level WARNING
        }
    }
    else {
        Write-ScriptLog -Message "No CDP or LLDP packets were received. Verify the network adapter is connected to a managed switch that has CDP or LLDP enabled." -LogFile $LogFile -Level WARNING
    }

    Write-ScriptLog -Message "Script finished." -LogFile $LogFile -Level INFO
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-ScriptLog -Message "CRITICAL ERROR: $ErrorMessage" -LogFile $LogFile -Level ERROR
    Write-ScriptLog -Message "Stack Trace: $($_.ScriptStackTrace)" -LogFile $LogFile -Level ERROR
    exit 1
}
