<#
.SYNOPSIS
    Retrieves and logs Azure Virtual Machine IMDS metadata.

.DESCRIPTION
    Queries the Azure Instance Metadata Service (IMDS) at http://169.254.169.254 to extract VM details 
    including VM Name, Size, Resource Group, Subscription ID, Region, Zone, IPs, Subnet, and MAC address.
    Outputs formatted summary to console and logs details to file.

.EXAMPLE
    .\Get-AzureVmDetails.ps1
#>

# Set log directory and timestamp formatting
$logDir = "$env:SystemDrive\Logs\Get-AzureVmDetails"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$imdsBaseUri = "http://169.254.169.254/metadata/instance?api-version=2021-02-01"

try {
    # Fetch full IMDS instance object
    $metadata = Invoke-RestMethod -Headers @{"Metadata"="true"} -Uri $imdsBaseUri -TimeoutSec 5
    $compute  = $metadata.compute
    $network  = $metadata.network.interface[0]

    # Parse key properties into a custom object
    $vmDetails = [ordered]@{
        "VM Name"           = $compute.name
        "VM Size"           = $compute.vmSize
        "OS Type"           = $compute.osType
        "Resource Group"    = $compute.resourceGroupName
        "Subscription ID"   = $compute.subscriptionId
        "Location / Region" = $compute.location
        "Zone"              = if ($compute.zone) { $compute.zone } else { "N/A" }
        "VM ID"             = $compute.vmId
        "Private IP"        = $network.ipv4.ipAddress[0].privateIpAddress
        "Public IP"         = if ($network.ipv4.ipAddress[0].publicIpAddress) { $network.ipv4.ipAddress[0].publicIpAddress } else { "None / Internal" }
        "Subnet Prefix"     = $network.ipv4.subnet[0].address + "/" + $network.ipv4.subnet[0].prefix
        "MAC Address"       = $network.macAddress
    }

    # Format output for console
    Write-Host "`n=== Azure VM IMDS Metadata Summary ===" -ForegroundColor Yellow
    $vmDetails.GetEnumerator() | ForEach-Object {
        Write-Host ("{0,-18} : {1}" -f $_.Key, $_.Value)
    }
    Write-Host ("=" * 38) "`n"

    # Export structured log entry using DDMMYY format
    $timestamp = Get-Date -Format "ddMMyy-HHmmss"
    $logPath   = Join-Path $logDir "AzureVM_Metadata_$timestamp.log"
    $vmDetails.GetEnumerator() | ForEach-Object { "$($_.Key) : $($_.Value)" } | Out-File -FilePath $logPath -Encoding utf8

    Write-Host "Log saved to: $logPath" -ForegroundColor DarkGray
}
catch {
    Write-Error "Failed to retrieve IMDS metadata. Ensure this script is running inside an Azure VM."
}
