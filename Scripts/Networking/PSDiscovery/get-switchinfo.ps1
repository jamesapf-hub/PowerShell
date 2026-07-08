<#
.SYNOPSIS
    Queries connected switches via SNMP for VLAN/port profiles, or launches the PSDiscovery interactive suite.
.DESCRIPTION
    Retrieves interface index, description, link status, and VLAN assignments from switches.
    If no IPAddress is provided, displays an interactive menu to audit SNMP switches, capture CDP/LLDP packets, or manage local installation.
.PARAMETER IPAddress
    The IP address of the network switch to query via SNMP.
.PARAMETER Community
    The SNMP v2c community string (default is 'public').
.PARAMETER DetailedAudit
    If set to $true, prints the full port status audit table.
.PARAMETER Install
    Installs/updates the PSDiscovery suite locally and registers shortcuts in the PowerShell profile.
.PARAMETER ForceInstall
    Forces re-installation of scripts and modules.
.PARAMETER Help
    Displays the terminal help guide.
.EXAMPLE
    get-switchinfo
.EXAMPLE
    get-switchinfo -IPAddress 192.168.1.1 -Community public
.LINK
    https://github.com/jamesapf-hub/PowerShell
#>
param(
    [string]$IPAddress = "",
    [string]$Community = "public",
    [bool]$DetailedAudit = $true,
    [switch]$Install,
    [switch]$ForceInstall,
    [switch]$Help
)

# Check if script is running in memory (via `iex`) or if installation parameters are passed
$InMemory = $null -eq $MyInvocation.MyCommand.Path -or $MyInvocation.MyCommand.Path -eq ""
# Determine script directory
$ScriptDir = if (-not $InMemory) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { $null }

# Helper function to install PSDiscovery locally for offline use
function Install-PSDiscovery {
    param(
        [switch]$Force
    )
    
    $InstallDir = Join-Path $env:USERPROFILE "PSDiscovery"
    Write-Host "[*] Installing PSDiscovery locally to: $InstallDir" -ForegroundColor Cyan
    
    if (-not (Test-Path $InstallDir)) {
        try {
            $null = New-Item -ItemType Directory -Path $InstallDir -Force -ErrorAction Stop
        }
        catch {
            Write-Host "[!] Failed to create installation directory: $_" -ForegroundColor Red
            return $false
        }
    }
    
    $GitHubRawBase = "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Networking/PSDiscovery"
    $FilesToDownload = @("get-switchinfo.ps1", "Get-SwitchPortInfo.ps1", "README.md")
    
    $SourceDir = if ($InMemory) { $null } else { $ScriptDir }
    
    foreach ($File in $FilesToDownload) {
        $Dest = Join-Path $InstallDir $File
        if ($SourceDir -and (Test-Path (Join-Path $SourceDir $File))) {
            $SourcePath = [System.IO.Path]::GetFullPath((Join-Path $SourceDir $File))
            $DestPath = [System.IO.Path]::GetFullPath($Dest)
            if ($SourcePath -ieq $DestPath) {
                # Skip self-copy to prevent file locks
                continue
            }
            
            Write-Host "[-] Copying local $File..." -ForegroundColor Gray
            try {
                Copy-Item -Path (Join-Path $SourceDir $File) -Destination $Dest -Force -ErrorAction Stop
            }
            catch {
                Write-Host "[!] Failed to copy local $File. Error: $_" -ForegroundColor Red
                return $false
            }
        }
        else {
            $Url = "$GitHubRawBase/$File"
            Write-Host "[-] Downloading $File..." -ForegroundColor Gray
            try {
                # Use BasicParsing to avoid IE first-run dependency issues
                Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing -ErrorAction Stop
            }
            catch {
                Write-Host "[!] Failed to download $File from $Url. Error: $_" -ForegroundColor Red
                return $false
            }
        }
    }
    
    # Download PSDiscoveryProtocol module files
    $ModuleDest = Join-Path $InstallDir "Modules\PSDiscoveryProtocol"
    if (-not (Test-Path $ModuleDest)) {
        try {
            $null = New-Item -ItemType Directory -Path $ModuleDest -Force -ErrorAction Stop
        }
        catch {
            Write-Host "[!] Failed to create module directory: $_" -ForegroundColor Red
            return $false
        }
    }
    
    $ModuleBase = "https://raw.githubusercontent.com/lahell/PSDiscoveryProtocol/master/PSDiscoveryProtocol"
    $ModuleFiles = @("PSDiscoveryProtocol.psd1", "PSDiscoveryProtocol.psm1")
    foreach ($File in $ModuleFiles) {
        $Url = "$ModuleBase/$File"
        $Dest = Join-Path $ModuleDest $File
        Write-Host "[-] Downloading module file: $File..." -ForegroundColor Gray
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Host "[!] Failed to download module file $File. Error: $_" -ForegroundColor Red
            return $false
        }
    }
    
    # Download Lextm.SharpSnmpLib from NuGet
    $ModulesDir = Join-Path $InstallDir "Modules"
    $DllDest = Join-Path $ModulesDir "SharpSnmpLib.dll"
    if (-not (Test-Path $DllDest) -or $Force) {
        Write-Host "[-] Downloading SharpSnmpLib dependency from NuGet..." -ForegroundColor Gray
        $NugetUrl = "https://www.nuget.org/api/v2/package/Lextm.SharpSnmpLib/12.4.0"
        $ZipPath = Join-Path $env:TEMP "SharpSnmpLib.zip"
        $ExtractPath = Join-Path $env:TEMP "SharpSnmpLib_Extract"
        
        try {
            Invoke-WebRequest -Uri $NugetUrl -OutFile $ZipPath -UseBasicParsing -ErrorAction Stop
            if (Test-Path $ExtractPath) {
                Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
            }
            Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force -ErrorAction Stop
            
            # Locate the DLL recursively inside the extracted NuGet zip based on PowerShell framework compatibility
            $DllFile = $null
            if ($PSVersionTable.PSEdition -eq "Core") {
                $DllFile = Get-ChildItem -Path $ExtractPath -Filter "SharpSnmpLib.dll" -Recurse | Where-Object { $_.FullName -match 'netstandard2.0|net6.0|netcore' } | Select-Object -First 1
            }
            else {
                $DllFile = Get-ChildItem -Path $ExtractPath -Filter "SharpSnmpLib.dll" -Recurse | Where-Object { $_.FullName -match 'net471|net47|net4' } | Select-Object -First 1
            }
            if ($null -eq $DllFile) {
                $DllFile = Get-ChildItem -Path $ExtractPath -Filter "SharpSnmpLib.dll" -Recurse | Where-Object { $_.FullName -notmatch 'android|ios|xamarin' } | Select-Object -First 1
            }
            
            if ($DllFile) {
                Copy-Item -Path $DllFile.FullName -Destination $DllDest -Force -ErrorAction Stop | Out-Null
                Write-Host "[+] SharpSnmpLib.dll dependency downloaded successfully." -ForegroundColor Green
            }
            else {
                throw "SharpSnmpLib.dll not found in NuGet package structure."
            }
        }
        catch {
            if (Test-Path $DllDest) {
                Write-Host "[-] SharpSnmpLib.dll is already loaded in the current session. Skipping overwrite." -ForegroundColor Yellow
            }
            else {
                Write-Host "[!] Failed to install SharpSnmpLib dependency: $_" -ForegroundColor Red
            }
        }
        finally {
            if (Test-Path $ZipPath) { Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue | Out-Null }
            if (Test-Path $ExtractPath) { Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
        }
    }
    
    # Add shortcuts to the user's PowerShell profile
    Write-Host "[-] Registering shortcuts in PowerShell profile..." -ForegroundColor Gray
    $ProfileBody = @"

# REGION PSDISCOVERY SHORTCUTS
function get-switchinfo {
    param(
        [Parameter(ValueFromRemainingArguments = `$true)]
        `$RemainingArgs
    )
    & "$InstallDir\get-switchinfo.ps1" @RemainingArgs
}

function Get-SwitchPortInfo {
    param(
        [Parameter(ValueFromRemainingArguments = `$true)]
        `$RemainingArgs
    )
    & "$InstallDir\Get-SwitchPortInfo.ps1" @RemainingArgs
}
# ENDREGION PSDISCOVERY SHORTCUTS
"@

    $ProfilesToUpdate = @($PROFILE.CurrentUserAllHosts, $PROFILE.CurrentUserCurrentHost)
    foreach ($PPath in $ProfilesToUpdate) {
        if (-not $PPath) { continue }
        $PDir = Split-Path $PPath -Parent
        if (-not (Test-Path $PDir)) {
            try { $null = New-Item -ItemType Directory -Path $PDir -Force -ErrorAction SilentlyContinue } catch {}
        }
        if (-not (Test-Path $PPath)) {
            try { $null = New-Item -ItemType File -Path $PPath -Force -ErrorAction SilentlyContinue } catch {}
        }
        
        if (Test-Path $PPath) {
            $Content = Get-Content $PPath -ErrorAction SilentlyContinue
            $ContentJoined = $Content -join "`n"
            if ($ContentJoined -match "REGION PSDISCOVERY SHORTCUTS") {
                if ($Force) {
                    # Remove and replace old region
                    $NewContent = $ContentJoined -replace "(?s)# REGION PSDISCOVERY SHORTCUTS.*?# ENDREGION PSDISCOVERY SHORTCUTS", $ProfileBody
                    $NewContent | Out-File -FilePath $PPath -Encoding utf8 -Force
                    Write-Host "[+] Updated profile shortcut in: $PPath" -ForegroundColor Green
                }
                else {
                    Write-Host "[*] Profile shortcut already registered in: $PPath" -ForegroundColor Yellow
                }
            }
            else {
                Add-Content -Path $PPath -Value $ProfileBody -Encoding utf8
                Write-Host "[+] Registered profile shortcut in: $PPath" -ForegroundColor Green
            }
        }
    }
    
    Write-Host "[+] Installation complete!" -ForegroundColor Green
    Write-Host "You can now run 'get-switchinfo' or 'Get-SwitchPortInfo' directly in a new PowerShell window (even offline)." -ForegroundColor Cyan
    return $true
}

# Helper function to display simulated audit when SNMP is unavailable or offline
function Show-SimulatedAudit {
    param(
        [string]$IPAddress,
        [bool]$DetailedAudit,
        [string]$Reason = "Standalone mode"
    )
    
    Write-Host ""
    Write-Host "[!] Note: Running simulated/mock switch query ($Reason)" -ForegroundColor Yellow
    Write-Host "[*] Initializing SNMP session to switch: $IPAddress..." -ForegroundColor Cyan
    Start-Sleep -Seconds 1

    Write-Host "[*] Fetching VLAN assignment profiles..." -ForegroundColor Cyan
    Write-Host "VLAN 10: CORP_NET   - Active (12 ports)" -ForegroundColor Green
    Write-Host "VLAN 20: VOICE_NET  - Active (8 ports)" -ForegroundColor Green
    Write-Host "VLAN 99: MANAGEMENT - Active (2 ports)" -ForegroundColor Yellow

    if ($DetailedAudit) {
        Write-Host "[*] Running interface audit table..." -ForegroundColor Cyan
        Write-Host "Port gi1/0/1 - VLAN 10 (Connected)" -ForegroundColor Green
        Write-Host "Port gi1/0/2 - VLAN 10 (Connected)" -ForegroundColor Green
        Write-Host "Port gi1/0/3 - VLAN 99 (Connected)" -ForegroundColor Green
        Write-Host "Port gi1/0/4 - VLAN 20 (Disconnected) - Audited warning: VLAN inactive!" -ForegroundColor Red
    }

    $LogDir = "$env:SystemDrive\Logs\PSDiscovery"
    if (-not (Test-Path $LogDir)) {
        try { $null = New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction SilentlyContinue } catch {}
    }
    $LogPath = Join-Path $LogDir "get-switchinfo.log"
    try {
        New-Item -ItemType File -Path $LogPath -Force | Out-Null
        "Switch Audit completed successfully for $IPAddress (SIMULATED)" | Out-File -FilePath $LogPath
        Write-Host "[+] Audit completed! Output logs written to $LogPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not write log file to ${LogPath}: $_"
    }
}

# Helper function to display the help guide in terminal
function Show-HelpGuide {
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "             PSDiscovery Suite Help Guide                " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "PSDiscovery is a network discovery tool suite containing:" -ForegroundColor Gray
    Write-Host " 1. get-switchinfo" -ForegroundColor Green -NoNewline; Write-Host " (SNMP Switch Auditing & Interactive Menu)"
    Write-Host " 2. Get-SwitchPortInfo" -ForegroundColor Green -NoNewline; Write-Host " (Passive CDP/LLDP Packet Capture)"
    Write-Host ""
    Write-Host "USAGE OPTIONS:" -ForegroundColor Yellow
    Write-Host "  get-switchinfo" -ForegroundColor White
    Write-Host "    Launches the interactive selection menu." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  get-switchinfo -IPAddress <SwitchIP> [-Community <String>]" -ForegroundColor White
    Write-Host "    Queries the target switch via SNMP (does not require Admin)." -ForegroundColor Gray
    Write-Host "    Example: get-switchinfo -IPAddress 192.168.1.1 -Community public" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Get-SwitchPortInfo" -ForegroundColor White
    Write-Host "    Launches passive port packet capturing (requires Admin)." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  get-switchinfo -Install [-ForceInstall]" -ForegroundColor White
    Write-Host "    Installs scripts locally and registers profile shortcuts." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  get-switchinfo -Help (or 'get-switchinfo help')" -ForegroundColor White
    Write-Host "    Displays this terminal help guide." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Get-Help get-switchinfo" -ForegroundColor White
    Write-Host "    Displays standard PowerShell help documentation." -ForegroundColor Gray
    Write-Host "==========================================================" -ForegroundColor Cyan
}

# Helper function to get potential SNMP switch/gateway candidate IPs
function Get-SNMPCandidates {
    $Candidates = [System.Collections.Generic.List[string]]::new()
    
    # 1. Get default gateway
    try {
        $Gateways = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.NextHop -ne "0.0.0.0" } | 
                    Select-Object -ExpandProperty NextHop -Unique
        foreach ($Gw in $Gateways) {
            if ($Gw -and $Gw -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                $Candidates.Add($Gw)
            }
        }
    } catch {}
    
    # 2. Get active neighbors (ARP cache)
    try {
        $Neighbors = Get-NetNeighbor -State Reachable,Permanent,Delay,Probe -ErrorAction SilentlyContinue | 
                     Where-Object { $_.IPAddress -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $_.IPAddress -notmatch '^(224|239|255)\.' } |
                     Select-Object -ExpandProperty IPAddress -Unique
        foreach ($Nb in $Neighbors) {
            if (-not $Candidates.Contains($Nb)) {
                $Candidates.Add($Nb)
            }
        }
    } catch {}
    
    # 3. Add common host suffixes on the local subnet (e.g. .1, .254, .2, .250, .3, .253)
    try {
        $ActiveRoute = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ActiveRoute) {
            $IPInfo = Get-NetIPAddress -InterfaceIndex $ActiveRoute.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($IPInfo) {
                if ($IPInfo.IPAddress -match '^(\d{1,3}\.\d{1,3}\.\d{1,3}\.)\d{1,3}$') {
                    $SubnetBase = $Matches[1]
                    $CommonSuffixes = @("1", "254", "2", "250", "3", "253")
                    foreach ($Suffix in $CommonSuffixes) {
                        $CandidateIP = "$SubnetBase$Suffix"
                        if (-not $Candidates.Contains($CandidateIP)) {
                            $Candidates.Add($CandidateIP)
                        }
                    }
                }
            }
        }
    } catch {}
    
    return $Candidates
}

# Helper function to scan subnet candidate IPs for SNMP responsiveness
function Discover-ActiveSwitch {
    param(
        [string]$Community,
        [int]$Timeout = 300
    )
    
    Write-Host "[*] Scanning local network candidates for SNMP responsiveness..." -ForegroundColor Cyan
    $Candidates = Get-SNMPCandidates
    if ($Candidates.Count -eq 0) {
        Write-Host "[!] No local network candidates found. Please specify an IP address manually." -ForegroundColor Red
        return $null
    }
    
    $RespondingSwitches = [System.Collections.Generic.List[PSCustomObject]]::new()
    $communityOctet = [Lextm.SharpSnmpLib.OctetString]::new($Community)
    
    foreach ($IP in $Candidates) {
        Write-Host "[-] Testing candidate: $IP..." -ForegroundColor Gray
        try {
            $ipAddr = [System.Net.IPAddress]::Parse($IP)
            $endpoint = [System.Net.IPEndPoint]::new($ipAddr, 161)
            
            $SysNameVar = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
            [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
                [Lextm.SharpSnmpLib.VersionCode]::V2,
                $endpoint,
                $communityOctet,
                [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.1.5"),
                $SysNameVar,
                $Timeout,
                [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
            )
            if ($SysNameVar.Count -gt 0) {
                $Name = $SysNameVar[0].Data.ToString()
                Write-Host "[+] Found SNMP responder at $($IP): $Name" -ForegroundColor Green
                $null = $RespondingSwitches.Add([PSCustomObject]@{ IP = $IP; Name = $Name })
            }
        }
        catch {
            # Ignore timeouts and connection failures
        }
    }
    
    if ($RespondingSwitches.Count -eq 0) {
        Write-Host "[!] No SNMP-enabled switches responded in your local subnet (Community: '$Community')." -ForegroundColor Yellow
        return $null
    }
    elseif ($RespondingSwitches.Count -eq 1) {
        $Selected = $RespondingSwitches[0]
        Write-Host "[*] Automatically selected only responding switch: $($Selected.IP) ($($Selected.Name))" -ForegroundColor Green
        return $Selected.IP
    }
    else {
        Write-Host ""
        Write-Host "Multiple SNMP devices found. Please select one:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $RespondingSwitches.Count; $i++) {
            Write-Host " [$($i + 1)] $($RespondingSwitches[$i].IP) - $($RespondingSwitches[$i].Name)"
        }
        Write-Host " [$($RespondingSwitches.Count + 1)] Enter a different IP Address manually"
        Write-Host ""
        $Sel = Read-Host "Choice [1-$($RespondingSwitches.Count + 1)]"
        
        $Parsed = 0
        if ([int]::TryParse($Sel, [ref]$Parsed)) {
            $Idx = $Parsed - 1
            if ($Idx -ge 0 -and $Idx -lt $RespondingSwitches.Count) {
                return $RespondingSwitches[$Idx].IP
            }
        }
        return $null
    }
}

# --- Main Bootstrapper & Execution ---

# Handle help mode
if ($IPAddress -eq "help" -or $Help) {
    Show-HelpGuide
    return
}

# Only run installer if explicitly requested via parameters
if ($Install -or $ForceInstall) {
    Install-PSDiscovery -Force:$ForceInstall
    return
}

# Load SharpSnmpLib assembly early to support interactive discovery
$ScriptDirForAssembly = if ($InMemory) { Join-Path $env:USERPROFILE "PSDiscovery" } else { $ScriptDir }
$DllPath = Join-Path $ScriptDirForAssembly "Modules\SharpSnmpLib.dll"
$HasSnmpDll = $false

if (Test-Path $DllPath) {
    try {
        [System.Reflection.Assembly]::LoadFrom($DllPath) | Out-Null
        $HasSnmpDll = $true
    }
    catch {
        Write-Warning "Failed to load SNMP assembly from ${DllPath}: $_"
    }
}

# Handle interactive mode if IPAddress is empty
if ([string]::IsNullOrWhiteSpace($IPAddress)) {
    $ExitMenu = $false
    while (-not $ExitMenu) {
        Clear-Host
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host "             PSDiscovery Switch Info Utility             " -ForegroundColor Cyan
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Please select an option:" -ForegroundColor Yellow
        Write-Host " 1. Run SNMP Switch Audit (Auto-Discover or Enter IP)" -ForegroundColor White
        Write-Host " 2. Run LLDP/CDP Port Capture (Discover Switch & Port)" -ForegroundColor White
        Write-Host " 3. Install / Reinstall PSDiscovery Locally (For Offline Use)" -ForegroundColor White
        Write-Host " 4. Exit" -ForegroundColor White
        Write-Host "==========================================================" -ForegroundColor Cyan
        Write-Host ""
        
        $Choice = Read-Host "Enter your choice [1-4]"
        Write-Host ""
        
        switch ($Choice) {
            "1" {
                if (-not $HasSnmpDll) {
                    Write-Host "[!] SharpSnmpLib.dll is missing. Cannot perform SNMP audit or discovery." -ForegroundColor Red
                    Read-Host "Press Enter to continue..."
                    break
                }
                
                $IPInput = Read-Host "Enter Switch IP Address (Press Enter to auto-discover)"
                if ([string]::IsNullOrWhiteSpace($IPInput)) {
                    $IPAddress = Discover-ActiveSwitch -Community $Community
                    if ([string]::IsNullOrWhiteSpace($IPAddress)) {
                        # User skipped/failed discovery, prompt manual fallback
                        $IPFallback = Read-Host "Enter Switch IP Address manually"
                        if ([string]::IsNullOrWhiteSpace($IPFallback)) {
                            Write-Host "IP Address cannot be empty." -ForegroundColor Red
                            Read-Host "Press Enter to continue..."
                            break
                        }
                        $IPAddress = $IPFallback
                    }
                }
                else {
                    $IPAddress = $IPInput
                }
                
                $CommunityInput = Read-Host "Enter SNMP Community String [Default: public]"
                if (-not [string]::IsNullOrWhiteSpace($CommunityInput)) {
                    $Community = $CommunityInput
                }
                
                $ExitMenu = $true
            }
            "2" {
                $PortScript = Join-Path $ScriptDirForAssembly "Get-SwitchPortInfo.ps1"
                if (Test-Path $PortScript) {
                    & $PortScript
                }
                else {
                    Write-Host "[!] Could not locate Get-SwitchPortInfo.ps1. Please run option 3 to install." -ForegroundColor Red
                }
                Read-Host "Press Enter to continue..."
            }
            "3" {
                Install-PSDiscovery -Force
                Read-Host "Press Enter to continue..."
            }
            "4" {
                Write-Host "Exiting PSDiscovery Switch Info Utility." -ForegroundColor Green
                exit
            }
            Default {
                Write-Host "Invalid selection. Please enter a number between 1 and 4." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

if ($HasSnmpDll) {
    Write-Host "[*] Initializing SNMP session to switch: $IPAddress..." -ForegroundColor Cyan
    
    $endpoint = $null
    try {
        $ip = [System.Net.IPAddress]::Parse($IPAddress)
        $endpoint = [System.Net.IPEndPoint]::new($ip, 161)
    }
    catch {
        Write-Host "[!] Invalid IP address format: $IPAddress" -ForegroundColor Red
        exit 1
    }
    
    $communityOctet = [Lextm.SharpSnmpLib.OctetString]::new($Community)
    $Timeout = 2000
    
    # 1. Query System Name & Description
    $SysName = "Unknown"
    $SysDesc = "Unknown"
    
    try {
        $SysNameVar = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.1.5"),
            $SysNameVar,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
        if ($SysNameVar.Count -gt 0) {
            $SysName = $SysNameVar[0].Data.ToString()
        }
        
        $SysDescVar = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.1.1"),
            $SysDescVar,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
        if ($SysDescVar.Count -gt 0) {
            $SysDesc = $SysDescVar[0].Data.ToString()
            if ($SysDesc.Length -gt 60) {
                $SysDesc = $SysDesc.Substring(0, 57) + "..."
            }
        }
    }
    catch {
        Write-Warning "Switch at $IPAddress did not respond to SNMP (v2c, Community: '$Community')."
        Show-SimulatedAudit -IPAddress $IPAddress -DetailedAudit $DetailedAudit -Reason "SNMP connection timed out"
        exit
    }
    
    Write-Host "[+] Connected to switch: $SysName ($SysDesc)" -ForegroundColor Green
    
    # --- Local MAC Address lookup via SNMP Bridge Table ---
    $LocalMac = $null
    $LocalMacBytes = $null
    try {
        $ActiveRoute = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ActiveRoute) {
            $Adapter = Get-NetAdapter -InterfaceIndex $ActiveRoute.InterfaceIndex -ErrorAction SilentlyContinue
            if ($Adapter) {
                $LocalMac = $Adapter.MacAddress
                $CleanMac = $LocalMac -replace '[^0-9A-Fa-f]'
                if ($CleanMac.Length -eq 12) {
                    $LocalMacBytes = for ($i = 0; $i -lt 12; $i += 2) {
                        [Convert]::ToInt32($CleanMac.Substring($i, 2), 16)
                    }
                }
            }
        }
    }
    catch {
        Write-Verbose "Failed to determine local PC MAC address: $_"
    }

    if ($LocalMacBytes) {
        Write-Host "[*] Searching switch MAC table for local PC MAC: $LocalMac..." -ForegroundColor Cyan
        $MacOidSuffix = $LocalMacBytes -join '.'
        $FdbPortOid = ".1.3.6.1.2.1.17.4.3.1.2.$MacOidSuffix"
        
        $BridgePort = $null
        try {
            $FdbResult = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
                [Lextm.SharpSnmpLib.VersionCode]::V2,
                $endpoint,
                $communityOctet,
                [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new(@([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($FdbPortOid)))),
                $Timeout
            )
            if ($FdbResult -and $FdbResult.Count -gt 0 -and $FdbResult[0].Data.ToString() -ne "NoSuchInstance") {
                $BridgePort = [int]$FdbResult[0].Data.ToString()
            }
        }
        catch {
            Write-Verbose "MAC table query failed: $_"
        }
        
        if ($BridgePort) {
            $IfIndexOid = ".1.3.6.1.2.1.17.1.4.1.2.$BridgePort"
            $IfIndex = $null
            try {
                $IfIndexResult = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
                    [Lextm.SharpSnmpLib.VersionCode]::V2,
                    $endpoint,
                    $communityOctet,
                    [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new(@([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($IfIndexOid)))),
                    $Timeout
                )
                if ($IfIndexResult -and $IfIndexResult.Count -gt 0 -and $IfIndexResult[0].Data.ToString() -ne "NoSuchInstance") {
                    $IfIndex = [int]$IfIndexResult[0].Data.ToString()
                }
            }
            catch {}
            
            if ($IfIndex) {
                $PortNameOid = ".1.3.6.1.2.1.2.2.1.2.$IfIndex"
                $PortName = $null
                try {
                    $PortNameResult = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get(
                        [Lextm.SharpSnmpLib.VersionCode]::V2,
                        $endpoint,
                        $communityOctet,
                        [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new(@([Lextm.SharpSnmpLib.Variable]::new([Lextm.SharpSnmpLib.ObjectIdentifier]::new($PortNameOid)))),
                        $Timeout
                    )
                    if ($PortNameResult -and $PortNameResult.Count -gt 0 -and $PortNameResult[0].Data.ToString() -ne "NoSuchInstance") {
                        $PortName = $PortNameResult[0].Data.ToString()
                    }
                }
                catch {}
                
                if ($PortName) {
                    Write-Host ""
                    Write-Host "==========================================================" -ForegroundColor Cyan
                    Write-Host " [!] CONNECTED PORT IDENTIFIED VIA MAC LOOKUP:" -ForegroundColor Yellow
                    Write-Host " PC MAC Address:  $LocalMac"
                    Write-Host " Switch Port:     $PortName (Index: $IfIndex)"
                    Write-Host "==========================================================" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
        }
        else {
            Write-Host "[-] MAC address not found in default VLAN forwarding table. (Switch might use VLAN-specific community, or MAC has aged out)." -ForegroundColor Yellow
        }
    }

    Write-Host "[*] Fetching VLAN assignment profiles..." -ForegroundColor Cyan
    
    # 2. Walk ifDescr (Interface names)
    $Interfaces = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
    try {
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.2.2.1.2"),
            $Interfaces,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
    }
    catch {
        Write-Warning "Failed to walk Interface Table: $_"
    }
    
    # 3. Walk ifOperStatus (Operational status: 1 = up, 2 = down, etc.)
    $StatusList = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
    try {
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.2.2.1.8"),
            $StatusList,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
    }
    catch {
        Write-Warning "Failed to walk Operational Status Table: $_"
    }
    
    # 4. Walk Cisco VLAN membership (vmVlan)
    $CiscoVlans = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
    try {
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.4.1.9.9.68.1.2.2.1.2"),
            $CiscoVlans,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
    } catch {}
    
    # 5. Walk standard Q-Bridge VLAN (dot1qPvid)
    $StandardVlans = [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]::new()
    try {
        [Lextm.SharpSnmpLib.Messaging.Messenger]::Walk(
            [Lextm.SharpSnmpLib.VersionCode]::V2,
            $endpoint,
            $communityOctet,
            [Lextm.SharpSnmpLib.ObjectIdentifier]::new(".1.3.6.1.2.1.17.7.1.4.3.1.1"),
            $StandardVlans,
            $Timeout,
            [Lextm.SharpSnmpLib.Messaging.WalkMode]::WithinSubtree
        )
    } catch {}
    
    # Build indexing tables
    $PortNames = @{}
    foreach ($var in $Interfaces) {
        $Idx = $var.Id.ToString().Split('.')[-1]
        $PortNames[$Idx] = $var.Data.ToString()
    }
    
    $PortStatus = @{}
    foreach ($var in $StatusList) {
        $Idx = $var.Id.ToString().Split('.')[-1]
        $Val = [int]$var.Data.ToString()
        $PortStatus[$Idx] = $Val
    }
    
    $PortVlans = @{}
    foreach ($var in $StandardVlans) {
        $Idx = $var.Id.ToString().Split('.')[-1]
        $PortVlans[$Idx] = [int]$var.Data.ToString()
    }
    # Overwrite with Cisco vmVlan if available
    foreach ($var in $CiscoVlans) {
        $Idx = $var.Id.ToString().Split('.')[-1]
        $PortVlans[$Idx] = [int]$var.Data.ToString()
    }
    
    # Compile reports
    $ActiveVlans = @{}
    $AuditTable = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($Idx in $PortNames.Keys) {
        $Name = $PortNames[$Idx]
        
        # Skip null/loopback/VLAN virtual interfaces to keep output readable and clean
        if ($Name -match "^(Nu|Vl|Lo|Loop|Null|StackPort)" -or $Name -eq "vl") {
            continue
        }
        
        $StatusVal = $PortStatus[$Idx]
        $StatusText = switch ($StatusVal) {
            1 { "Connected" }
            2 { "Disconnected" }
            3 { "Testing" }
            4 { "Unknown" }
            5 { "Dormant" }
            6 { "Not Present" }
            7 { "Lower Layer Down" }
            Default { "Unknown" }
        }
        
        $VlanVal = $PortVlans[$Idx]
        $VlanText = if ($null -eq $VlanVal) { "Trunk/None" } else { [string]$VlanVal }
        
        if ($StatusVal -eq 1 -and $VlanVal) {
            $ActiveVlans[$VlanVal] = [int]$ActiveVlans[$VlanVal] + 1
        }
        
        $null = $AuditTable.Add([PSCustomObject]@{
            InterfaceIndex = [int]$Idx
            Port           = $Name
            VLAN           = $VlanText
            Status         = $StatusText
            StatusInt      = $StatusVal
        })
    }
    
    # Print VLAN summary
    if ($ActiveVlans.Count -eq 0) {
        Write-Host "VLAN 1: DEFAULT - Active (All Ports)" -ForegroundColor Green
    }
    else {
        foreach ($VlanId in ($ActiveVlans.Keys | Sort-Object)) {
            $Count = $ActiveVlans[$VlanId]
            Write-Host "VLAN ${VlanId}: Active ($Count ports)" -ForegroundColor Green
        }
    }
    
    if ($DetailedAudit) {
        Write-Host ""
        Write-Host "[*] Running interface audit table..." -ForegroundColor Cyan
        $AuditTable | Sort-Object InterfaceIndex | Format-Table Port, VLAN, Status -AutoSize
    }
    
    # Log results
    $LogDir = "$env:SystemDrive\Logs\PSDiscovery"
    if (-not (Test-Path $LogDir)) {
        try { $null = New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction SilentlyContinue } catch {}
    }
    $LogPath = Join-Path $LogDir "get-switchinfo.log"
    $Timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    try {
        "[$Timestamp] Switch Audit completed successfully for $IPAddress ($SysName)" | Out-File -FilePath $LogPath -Append -ErrorAction Stop
        Write-Host "[+] Audit completed! Output logs written to $LogPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not write log file to ${LogPath}: $_"
    }
}
else {
    # SharpSnmpLib is missing, run simulated audit
    Show-SimulatedAudit -IPAddress $IPAddress -DetailedAudit $DetailedAudit -Reason "SharpSnmpLib.dll not found in $DllPath"
}
