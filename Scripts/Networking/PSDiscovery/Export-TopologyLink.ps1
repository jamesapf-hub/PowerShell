<#
.SYNOPSIS
    Generates a HelperPage Network Topology deep link from PowerShell objects.
.DESCRIPTION
    Takes an array of network device objects, constructs the Topology DSL string, Base64 encodes it, 
    and outputs a clickable URL that automatically loads the diagram in your portal.
.PARAMETER Devices
    An array of objects with properties like Name, Type, IP, and ConnectedTo.
.EXAMPLE
    $devices = @(
        [PSCustomObject]@{ Name="Core-Switch"; Type="switch"; IP="192.168.1.1" },
        [PSCustomObject]@{ Name="AP-1"; Type="ap"; ConnectedTo="Core-Switch" },
        [PSCustomObject]@{ Name="AP-2"; Type="ap"; ConnectedTo="Core-Switch" }
    )
    $devices | Export-TopologyLink
#>
function Export-TopologyLink {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        [PSObject[]]$Devices
    )

    begin {
        $Script:DslLines = New-Object System.Collections.Generic.List[string]
        $Script:DeviceList = New-Object System.Collections.Generic.List[PSObject]
    }

    process {
        foreach ($Device in $Devices) {
            $Script:DeviceList.Add($Device)
        }
    }

    end {
        # 1. Generate node definitions
        foreach ($Device in $Script:DeviceList) {
            $NodeString = $Device.Name
            
            $Attributes = New-Object System.Collections.Generic.List[string]
            if ($Device.Type) { $Attributes.Add("type: $($Device.Type)") }
            if ($Device.IP) { $Attributes.Add("ip: $($Device.IP)") }
            if ($Device.Model) { $Attributes.Add("model: $($Device.Model)") }
            if ($Device.Brand) { $Attributes.Add("brand: $($Device.Brand)") }

            if ($Attributes.Count -gt 0) {
                $NodeString += " [" + ($Attributes -join ", ") + "]"
            }
            $Script:DslLines.Add($NodeString)
        }

        # 2. Generate edge definitions
        foreach ($Device in $Script:DeviceList) {
            if ($Device.ConnectedTo) {
                $Script:DslLines.Add("$($Device.ConnectedTo) > $($Device.Name)")
            }
        }

        $DslString = $Script:DslLines -join "`n"
        
        # 3. Base64 encode the DSL string
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($DslString)
        $Base64 = [System.Convert]::ToBase64String($Bytes)

        # 4. Generate the URL
        $Url = "https://helperpage.jamesapf.workers.dev/utilities/networktopology#dsl=$Base64"
        
        Write-Host ""
        Write-Host "[+] Network Topology DSL successfully generated!" -ForegroundColor Green
        Write-Host "Hold CTRL and click the link below to view it:" -ForegroundColor Cyan
        Write-Host $Url -ForegroundColor White
        Write-Host ""
        
        return $Url
    }
}
