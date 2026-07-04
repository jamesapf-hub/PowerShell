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
$ModuleImported = $false

if ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $CandidatePaths = @(
        (Join-Path $ScriptDir "..\Modules\Logging\Logging.psm1"),
        (Join-Path $ScriptDir "Modules\Logging\Logging.psm1"),
        (Join-Path $ScriptDir "..\..\Modules\Logging\Logging.psm1"),
        (Join-Path $ScriptDir "..\..\..\Modules\Logging\Logging.psm1")
    )

    foreach ($Path in $CandidatePaths) {
        if (Test-Path $Path) {
            Import-Module $Path -Force -ErrorAction SilentlyContinue
            $ModuleImported = $true
            break
        }
    }
}

if (-not $ModuleImported) {
    # Define fallback logging functions so the script can run standalone/remotely (e.g., from raw GitHub URL)
    function Initialize-ScriptLogging {
        param (
            [string]$ScriptPath,
            [string]$LogDirectory
        )
        if ([string]::IsNullOrEmpty($LogDirectory)) {
            $SystemDrive = $env:SystemDrive
            if ([string]::IsNullOrEmpty($SystemDrive)) {
                $SystemDrive = "C:"
            }
            $LogDirectory = Join-Path $SystemDrive "Log\PSDiscovery"
        }
        if (-not (Test-Path -Path $LogDirectory -PathType Container)) {
            try {
                $null = New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue
            } catch {}
        }
        $DateString = Get-Date -Format "dd-MM-yy"
        return Join-Path $LogDirectory "Get-SwitchPortInfo_${DateString}.log"
    }

    function Write-ScriptLog {
        param (
            [string]$Message,
            [string]$LogFile,
            [string]$Level = 'INFO'
        )
        $Timestamp = Get-Date -Format "dd\/MM\/yy HH:mm:ss"
        $LogEntry = "[$Timestamp] [$Level] $Message"
        switch ($Level) {
            'ERROR'   { Write-Host $LogEntry -ForegroundColor Red }
            'WARNING' { Write-Host $LogEntry -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $LogEntry -ForegroundColor Green }
            Default   { Write-Host $LogEntry -ForegroundColor White }
        }
        try {
            Add-Content -Path $LogFile -Value $LogEntry -ErrorAction SilentlyContinue
        } catch {}
    }
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
    Write-ScriptLog -Message "Checking for PSDiscoveryProtocol module dependency..." -LogFile $LogFile -Level INFO
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
    else {
        Write-ScriptLog -Message "PSDiscoveryProtocol module already installed." -LogFile $LogFile -Level SUCCESS
    }

    # Explicitly import the module to ensure all cmdlets are available
    Import-Module -Name PSDiscoveryProtocol -Force -ErrorAction Stop

    # 4. Interactive Menu Loop
    $ExitMenu = $false
    $FirstRun = $true
    while (-not $ExitMenu) {
        if (-not $FirstRun) {
            Clear-Host
        }
        $FirstRun = $false
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host "             PSDiscovery Switch Utility                  " -ForegroundColor Cyan
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host " 1. Capture/Update Switch Port Info (CDP/LLDP)" -ForegroundColor White
        Write-Host " 2. View Switch Inventory CSV File" -ForegroundColor White
        Write-Host " 3. View Switch Discovery Logs" -ForegroundColor White
        Write-Host " 4. Exit" -ForegroundColor White
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host ""
        
        $Selection = Read-Host "Enter your choice [1-4]"
        Write-Host ""
        
        switch ($Selection) {
            "1" {
                Write-ScriptLog -Message "Starting packet capture..." -LogFile $LogFile -Level INFO
                Write-ScriptLog -Message "Listening for CDP/LLDP packets on network interfaces (this may take up to 60 seconds)..." -LogFile $LogFile -Level INFO

                try {
                    $Packet = Invoke-DiscoveryProtocolCapture -ErrorAction Stop
                }
                catch {
                    Write-ScriptLog -Message "Capture command failed. Make sure you are running as Administrator or have WinPcap/Npcap installed if required." -LogFile $LogFile -Level ERROR
                    $Packet = $null
                }

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

                        # Export/Append findings to the CSV file
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
            }
            "2" {
                if (Test-Path $CsvPath) {
                    Write-Host "--- Switch Inventory CSV (Path: $CsvPath) ---" -ForegroundColor Yellow
                    Import-Csv -Path $CsvPath | Format-Table -AutoSize
                }
                else {
                    Write-Host "No Switch Inventory CSV found at $CsvPath." -ForegroundColor Red
                    Write-Host "Please run the discovery first (Option 1)." -ForegroundColor Red
                }
            }
            "3" {
                if (Test-Path $LogFile) {
                    Write-Host "--- Recent Log Entries (Path: $LogFile) ---" -ForegroundColor Yellow
                    Get-Content -Path $LogFile -Tail 20
                }
                else {
                    Write-Host "Log file not found at $LogFile." -ForegroundColor Red
                }
            }
            "4" {
                $ExitMenu = $true
                Write-ScriptLog -Message "User exited the interactive menu." -LogFile $LogFile -Level INFO
                Write-Host "Thank you for using PSDiscovery Switch Utility!" -ForegroundColor Green
            }
            Default {
                Write-Host "Invalid selection. Please enter a number between 1 and 4." -ForegroundColor Red
            }
        }
        
        if (-not $ExitMenu) {
            Write-Host ""
            Read-Host "Press Enter to continue..."
        }
    }
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-ScriptLog -Message "CRITICAL ERROR: $ErrorMessage" -LogFile $LogFile -Level ERROR
    Write-ScriptLog -Message "Stack Trace: $($_.ScriptStackTrace)" -LogFile $LogFile -Level ERROR
    exit 1
}
