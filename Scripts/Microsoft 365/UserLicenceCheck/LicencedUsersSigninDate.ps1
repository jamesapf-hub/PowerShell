<#
.SYNOPSIS
    Retrieves all licensed Entra ID / M365 users alongside their assigned licenses and last successful sign-in date.
.DESCRIPTION
    Queries Microsoft Graph for users with assigned licenses, maps the SkuIds to human-readable names 
    using local dictionary fallbacks, and extracts sign-in details.
.PARAMETER None
.NOTES
    Author:  JP
    Version: 2.0.0
    Date:    260624
#>

# Define current date for logs and filenames
$CurrentDate = (Get-Date).ToString("ddMMyy")

# 1. Setup Log Directory (fixed location C:\Seriun\log)
$LogDirectory = "C:\Seriun\log"
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$LogFile = Join-Path $LogDirectory "LicensedUsers_RunLog_${CurrentDate}.log"

# Helper function to write to both Host and Log file
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Type = 'Info',
        [ConsoleColor]$ForegroundColor = 'Cyan'
    )
    $Timestamp = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
    $LogEntry = "[$Timestamp] [$Type] $Message"
    
    # Write to console
    $ConsoleColor = switch ($Type) {
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        default { $ForegroundColor }
    }
    
    if ($Type -eq 'Warning') {
        Write-Warning $Message
    } elseif ($Type -eq 'Error') {
        Write-Error $Message
    } else {
        Write-Host $Message -ForegroundColor $ConsoleColor
    }
    
    # Write to log file
    $LogEntry | Out-File -FilePath $LogFile -Append -Encoding utf8
}

# Start logging execution
Write-Log "--------------------------------------------------" -ForegroundColor DarkGray
Write-Log "Execution started." -ForegroundColor Green

# Check for Microsoft.Graph module and install if missing
$RequiredModule = "Microsoft.Graph"
Write-Log "Checking if $RequiredModule module is installed..."
$ModuleCheck = Get-Module -ListAvailable -Name $RequiredModule
if (-not $ModuleCheck) {
    Write-Log "$RequiredModule module is missing. Attempting automatic installation for current user..." -Type Warning
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Log "Installing $RequiredModule module from PSGallery. This may take a few minutes..."
        Install-Module -Name $RequiredModule -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        Write-Log "$RequiredModule module successfully installed." -Type Success
    } catch {
        Write-Log "Failed to install $RequiredModule module automatically. Error: $_" -Type Error
        Write-Log "Please install it manually by running: Install-Module $RequiredModule -Scope CurrentUser" -Type Warning
        Write-Log "Execution halted due to missing dependencies." -ForegroundColor Red
        return
    }
} else {
    Write-Log "$RequiredModule module is already installed." -Type Success
}

# 2. Setup Output Directory (folder in the location the script is run from)
$ScriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$OutputDirectory = Join-Path $ScriptDirectory "Output"
if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    Write-Log "Created output directory: $OutputDirectory" -ForegroundColor Gray
} else {
    Write-Log "Using output directory: $OutputDirectory" -ForegroundColor Gray
}

# Define Clear-MgGraphCache function
function Clear-MgGraphCache {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try {
        Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path "$env:LOCALAPPDATA\.IdentityService") {
            Get-ChildItem "$env:LOCALAPPDATA\.IdentityService" -Filter "mg*" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        }
    } catch {}
}

# Connect to Microsoft Graph (disconnect first to ensure a clean start)
Write-Log "Clearing Microsoft Graph cache for a clean start..."
Clear-MgGraphCache
Write-Log "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Organization.Read.All", "LicenseAssignment.Read.All" -ContextScope Process

# Dynamically retrieve Tenant Display Name
Write-Log "Fetching tenant metadata..."
try {
    $OrgInfo = Get-MgOrganization | Select-Object -First 1 -Property DisplayName
    $TenantName = $OrgInfo.DisplayName -replace '[\\/:*?"<>| ]', '_'
} catch {
    Write-Log "Could not retrieve tenant name. Defaulting to 'UnknownTenant'." -Type Warning
    $TenantName = "UnknownTenant"
}

$OutputFile = Join-Path $OutputDirectory "LicensedUsers_${TenantName}_${CurrentDate}.csv"

# Build a Master SKU Reference Mapping Table
# 2. Offline Master Dictionary Mapping for missing/trial/free/add-on SKUs
$LocalSkuDictionary = @{
    # New additions from image highlights
    "ab5128ae-2475-4d95-8c73-33f07d701bfc" = "Microsoft 365 Copilot"
    "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Power BI Pro"
    "639dec6b-bb19-468b-871c-c5c441c4b0cb" = "Microsoft 365 Copilot"
    "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
    "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"

    # Tenant specific additions
    "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
    "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
    "61346032-1554-4736-b876-7d28d697f394" = "Microsoft 365 F3"
    "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87" = "Microsoft 365 F3"
    "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
    "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
    
    # Common trial, viral, and core enterprise plans
    "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "Microsoft Teams Exploratory"
    "cbdc14ab-d96c-4c30-b9f4-6bc7c7bd7d72" = "Microsoft 365 Business Premium"
    "bd251394-b1ed-487b-a1aa-ee198c62c938" = "Microsoft 365 Business Basic"
    "078d10ee-6995-4851-8043-334f610f49b3" = "Microsoft 365 E5"
    "47d12459-1159-4ad1-abfa-00e92056813a" = "Microsoft 365 E3"
    "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
    "c5928f49-12ba-48f7-ada3-0d743a3dbd2b" = "Office 365 E5"
    "a403a5cc-140d-4168-bfb5-58b5e7cbf094" = "Microsoft 365 Audio Conferencing"
    "57ff2da0-73d6-437c-a4df-ab9dc44faefc" = "Microsoft Service Business"

    # Technical Name Overrides
    "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
    "BUSINESS_PREMIUM"         = "Microsoft 365 Business Premium"
    "SPB"                      = "Microsoft 365 Business Premium"
    "SMB_BUSINESS"             = "Microsoft 365 Business Premium"
    "EXCHANGESTANDARD"         = "Exchange Online (Plan 1)"
    "AAD_PREMIUM"              = "Microsoft Entra ID P1"
    "POWERAPPS_DEV"            = "Microsoft Power Apps for Developer"
    "TEAMS_ROOM_BASIC"         = "Microsoft Teams Rooms Basic"
    "TEAMS_ROOM_PRO"           = "Microsoft Teams Rooms Pro"
}

$SkuMap = @{}
# Pre-populate friendly names from local dictionary
foreach ($Id in $LocalSkuDictionary.Keys) {
    $SkuMap[$Id.ToLower()] = $LocalSkuDictionary[$Id] -replace ';', ','
}

# Pre-populate friendly names from LicensePrices.csv
$PricesFile = Join-Path $ScriptDirectory "LicensePrices.csv"
if (Test-Path $PricesFile) {
    try {
        $PricesCsv = Import-Csv $PricesFile
        foreach ($row in $PricesCsv) {
            if ($row.SkuId -and $row.Product) {
                $SkuMap[$row.SkuId.ToString().ToLower()] = $row.Product
            }
        }
    } catch {}
}

# 1. Fetch live tenant SKUs first
Write-Log "Fetching tenant local license SKUs..."
try {
    if ($null -eq $Script:SkuPrices) {
        $Script:SkuPrices = Initialize-SkuPrices
    }
    
    $SubSkus = Get-MgSubscribedSku -All
    foreach ($sku in $SubSkus) {
        if ($sku.SkuId -and $sku.SkuPartName) {
            $GuidStr = $sku.SkuId.ToString().ToLower()
            if (-not $SkuMap.ContainsKey($GuidStr)) {
                $SkuMap[$GuidStr] = $sku.SkuPartName -replace ';', ','
            }
            
            # Resolve SKU Name using GUID mapping, then SkuPartName override mapping, then raw part name
            $SkuName = if ($SkuMap.ContainsKey($GuidStr)) { 
                $SkuMap[$GuidStr] 
            } elseif ($SkuMap.ContainsKey($sku.SkuPartName)) { 
                $SkuMap[$sku.SkuPartName] 
            } else { 
                $sku.SkuPartName -replace ';', ',' 
            }
            
            # Fallback for empty names
            if ([string]::IsNullOrWhiteSpace($SkuName)) {
                $SkuName = "Unknown SKU ($GuidStr)"
            }
            
            # If SkuId is not in the CSV pricing database, estimate and append it
            if ($null -ne $Script:SkuPrices -and -not $Script:SkuPrices.ContainsKey($GuidStr)) {
                $UnitPrice = Get-LicenseMonthlyPrice $SkuName
                if ($SkuName -notlike "Unknown SKU*" -and -not [string]::IsNullOrWhiteSpace($SkuName)) {
                    try {
                        Add-Content -Path $PricesFile -Value "`n$SkuName,$UnitPrice,$GuidStr" -Encoding utf8
                        $Script:SkuPrices[$GuidStr] = $UnitPrice
                        $Script:SkuPrices[$SkuName] = $UnitPrice
                        Write-Log "Discovered new SKU '$SkuName' ($GuidStr). Added to LicensePrices.csv with estimate £{0:N2}." -f $UnitPrice
                    } catch {}
                }
            }
        }
    }
} catch {
    Write-Log "Failed to query local tenant SKUs. Error: $_" -Type Warning
}

# 3. Dynamic Global SKU Reference Loading (with local caching to speed up execution)
$CacheFile = Join-Path $LogDirectory "M365_SKU_Cache.csv"
$CacheExpirationDays = 7
$LoadedGlobalSkus = $false

# Check if a fresh cache exists
if (Test-Path $CacheFile) {
    $LastModified = (Get-Item $CacheFile).LastWriteTime
    if ($LastModified -gt (Get-Date).AddDays(-$CacheExpirationDays)) {
        Write-Log "Loading global SKU map from local cache..."
        try {
            $CachedSkus = Import-Csv -Path $CacheFile
            foreach ($Sku in $CachedSkus) {
                if ($Sku.GUID -and $Sku.Product_Display_Name) {
                    $Guid = $Sku.GUID.ToString().ToLower()
                    if (-not $SkuMap.ContainsKey($Guid)) {
                        $SkuMap[$Guid] = $Sku.Product_Display_Name -replace ';', ','
                    }
                }
            }
            $LoadedGlobalSkus = $true
        } catch {
            Write-Log "Failed to load cache. Re-fetching online..." -Type Warning
        }
    }
}

# If no fresh cache is available, fetch online
if (-not $LoadedGlobalSkus) {
    Write-Log "Fetching latest global SKU mapping database online..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # Fetch community-updated license mapping database
        $OnlineSkuCsvText = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/merill/license/main/license.csv" -TimeoutSec 10
        
        # Parse and load into SkuMap
        $OnlineSkus = ConvertFrom-Csv -InputObject $OnlineSkuCsvText
        foreach ($Sku in $OnlineSkus) {
            if ($Sku.GUID -and $Sku.Product_Display_Name) {
                $Guid = $Sku.GUID.ToString().ToLower()
                if (-not $SkuMap.ContainsKey($Guid)) {
                    $SkuMap[$Guid] = $Sku.Product_Display_Name -replace ';', ','
                }
            }
        }
        
        # Save fresh copy to cache folder
        $OnlineSkuCsvText | Out-File -FilePath $CacheFile -Force -Encoding utf8
        Write-Log "Successfully updated local SKU cache database." -Type Success
    } catch {
        Write-Log "Could not reach online SKU repository. Using built-in local default fallbacks only." -Type Warning
    }
}

# Filter only for users that have at least one assigned license
Write-Log "Retrieving licensed users and sign-in activity..."

# Disable progress bar rendering temporarily to speed up large query execution
$OriginalProgressPreference = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'

try {
    $UserProperties = @('Id', 'DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'SignInActivity')
    $LicensedUsers = Get-MgUser -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable LicensedCount -All -Property $UserProperties
} catch {
    Write-Log "Failed to retrieve users from Microsoft Graph. Error: $_" -Type Error
    $ProgressPreference = $OriginalProgressPreference
    Write-Log "Execution halted due to query failure." -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    return
} finally {
    # Restore original progress preference
    $ProgressPreference = $OriginalProgressPreference
}

function Initialize-SkuPrices {
    $PricesFile = Join-Path $ScriptDirectory "LicensePrices.csv"
    
    # Auto-generate default CSV if missing
    if (-not (Test-Path $PricesFile)) {
        $DefaultCsv = @"
Product,MonthlyPrice,SkuId
Microsoft 365 Business Premium,18.10,cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46
Microsoft 365 Business Standard,10.30,f245ecc8-75af-4f8e-b61f-27d8114de5f3
Microsoft 365 Business Basic,4.90,bd251394-b1ed-487b-a1aa-ee198c62c938
Microsoft 365 Business Basic,4.90,3b555118-da6a-4418-894f-7df1e2096870
Microsoft 365 E5,48.10,078d10ee-6995-4851-8043-334f610f49b3
Microsoft 365 E3,31.70,47d12459-1159-4ad1-abfa-00e92056813a
Microsoft 365 F3,6.60,61346032-1554-4736-b876-7d28d697f394
Microsoft 365 F3,6.60,f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87
Exchange Online (Plan 2),6.60,19ec0d23-8335-4cbd-94ac-6050e30712fa
Power BI Pro,8.20,a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
Microsoft 365 Copilot,24.70,639dec6b-bb19-468b-871c-c5c441c4b0cb
Microsoft 365 Copilot,24.70,ab5128ae-2475-4d95-8c73-33f07d701bfc
Office 365 E3,22.00,6fd2c87f-b296-42f0-b197-1e91e994b900
Office 365 E5,37.50,c5928f49-12ba-48f7-ada3-0d743a3dbd2b
Microsoft 365 Audio Conferencing,3.30,a403a5cc-140d-4168-bfb5-58b5e7cbf094
Microsoft Service Business,0.00,57ff2da0-73d6-437c-a4df-ab9dc44faefc
Microsoft Teams Exploratory,0.00,710779e8-3d4a-4c88-adb9-386c958d1fdf
Microsoft Power Automate Free,0.00,f30db892-07e9-47e9-837c-80727f46fd3d
Windows Store for Business,0.00,6470687e-a428-4b7a-bef2-8a291ad947c9
"@
        try {
            $DefaultCsv | Out-File -FilePath $PricesFile -Force -Encoding utf8
        } catch {
            Write-Log "Warning: Could not create default LicensePrices.csv: $_" -Type Warning
        }
    } else {
        # Check if user's existing CSV is missing critical GUIDs and append them
        try {
            $CsvContent = Get-Content -Path $PricesFile -Raw
            if ($CsvContent -notmatch "ab5128ae-2475-4d95-8c73-33f07d701bfc") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 Copilot,24.70,ab5128ae-2475-4d95-8c73-33f07d701bfc" -Encoding utf8
            }
            if ($CsvContent -notmatch "61346032-1554-4736-b876-7d28d697f394") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 F3,6.60,61346032-1554-4736-b876-7d28d697f394" -Encoding utf8
            }
            if ($CsvContent -notmatch "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 F3,6.60,f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87" -Encoding utf8
            }
            if ($CsvContent -notmatch "6470687e-a428-4b7a-bef2-8a291ad947c9") {
                Add-Content -Path $PricesFile -Value "`nWindows Store for Business,0.00,6470687e-a428-4b7a-bef2-8a291ad947c9" -Encoding utf8
            }
        } catch {}

        # Self-healing cleanup of any malformed blank entries or legacy incorrect mappings (e.g. 3b555118 F3 mapping)
        try {
            $CleanLines = [System.Collections.Generic.List[string]]::new()
            $HasMalformed = $false
            foreach ($line in Get-Content -Path $PricesFile) {
                if ($line -like "Product,MonthlyPrice,SkuId") {
                    $CleanLines.Add($line)
                    continue
                }
                # Fix F3 incorrect mapping to Business Basic
                if ($line -like "*3b555118-da6a-4418-894f-7df1e2096870*") {
                    $CleanLines.Add("Microsoft 365 Business Basic,4.90,3b555118-da6a-4418-894f-7df1e2096870")
                    $HasMalformed = $true
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($line) -or $line -like ",*") {
                    $HasMalformed = $true
                    continue
                }
                $CleanLines.Add($line)
            }
            if ($HasMalformed) {
                $CleanLines | Out-File -FilePath $PricesFile -Force -Encoding utf8
            }
        } catch {}
    }
    
    $Prices = @{}
    if (Test-Path $PricesFile) {
        try {
            $Csv = Import-Csv -Path $PricesFile
            foreach ($row in $Csv) {
                $Price = 0.00
                if ($row.MonthlyPrice -and [double]::TryParse($row.MonthlyPrice, [ref]$Price)) {
                    # Map by Product Name
                    if ($row.Product) {
                        $Prices[$row.Product.Trim()] = $Price
                    }
                    # Map by SkuId (GUID) as well to allow exact GUID lookups!
                    if ($row.SkuId -and $row.SkuId.Trim() -ne "") {
                        $Guid = $row.SkuId.Trim().ToLower()
                        $Prices[$Guid] = $Price
                    }
                }
            }
        } catch {
            Write-Log "Error loading LicensePrices.csv: $_" -Type Warning
        }
    }
    
    # If loading failed or file was empty, fall back to inline defaults
    if ($Prices.Count -eq 0) {
        $Prices = @{
            "Microsoft 365 Business Premium"     = 18.10
            "Microsoft 365 Business Standard"    = 10.30
            "Microsoft 365 Business Basic"       = 4.90
            "Microsoft 365 E5"                   = 48.10
            "Microsoft 365 E3"                   = 31.70
            "Microsoft 365 F3"                   = 6.60
            "Exchange Online (Plan 2)"           = 6.60
            "Power BI Pro"                       = 8.20
            "Microsoft 365 Copilot"              = 24.70
            "Office 365 E3"                      = 22.00
            "Office 365 E5"                      = 37.50
            "Microsoft 365 Audio Conferencing"   = 3.30
            "Microsoft Service Business"         = 0.00
            "Microsoft Teams Exploratory"        = 0.00
            "Microsoft Power Automate Free"      = 0.00
        }
    }
    
    return $Prices
}

$SkuPrices = Initialize-SkuPrices

function Get-LicenseMonthlyPrice($LicensesString) {
    if (-not $LicensesString) { return 0.00 }
    $Sum = 0.00
    $LicensesString.Split(',') | ForEach-Object {
        $Name = $_.Trim()
        if ($Name) {
            # Extract GUID if the name is formatted like "Unknown SKU (guid)"
            if ($Name -match "Unknown SKU \(([^)]+)\)") {
                $Name = $Matches[1].Trim().ToLower()
            }
            
            $Matched = $false
            foreach ($key in $SkuPrices.Keys) {
                # Check for exact display name, start/end matches, or exact SkuId GUID match
                if ($Name -eq $key -or $Name.StartsWith($key) -or $key.StartsWith($Name) -or ($Name -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$" -and $Name -eq $key.ToLower())) {
                    $Sum += $SkuPrices[$key]
                    $Matched = $true
                    break
                }
            }
            if (-not $Matched) {
                if ($Name -match "E5") { $Sum += 48.10 }
                elseif ($Name -match "E3") { $Sum += 31.70 }
                elseif ($Name -match "Business") { $Sum += 10.30 }
                else { $Sum += 10.00 }
            }
        }
    }
    return $Sum
}

function Parse-DateString($DateStr) {
    if ($null -eq $DateStr) { return $null }
    if ($DateStr -is [System.DateTime]) { return $DateStr }
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $null }
    if ($DateStr -like "*no interactive*" -or $DateStr -like "*recorded*") { return $null }
    
    # Try en-GB (UK) first
    [DateTime]$ParsedDate = [DateTime]::MinValue
    $UkCulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-GB")
    if ([DateTime]::TryParse([string]$DateStr, $UkCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    
    # Try en-US second
    $UsCulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
    if ([DateTime]::TryParse([string]$DateStr, $UsCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    
    # Fallback to current culture TryParse
    if ([DateTime]::TryParse([string]$DateStr, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    return $null
}

function Get-DaysSince($DateStr) {
    $Date = Parse-DateString $DateStr
    if ($null -eq $Date) { return [double]::PositiveInfinity }
    
    $Diff = (Get-Date) - $Date
    if ($Diff.TotalDays -lt 0) { return 0 }
    return [Math]::Floor($Diff.TotalDays)
}

$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($User in $LicensedUsers) {
    # Resolve friendly names utilizing the built master map
    $UserLicenses = foreach ($License in $User.AssignedLicenses) {
        $TargetGuid = $License.SkuId.ToString().ToLower()
        if ($SkuMap.ContainsKey($TargetGuid)) {
            $SkuMap[$TargetGuid]
        } else {
            "Unknown SKU ($TargetGuid)"
        }
    }
    $LicenseString = $UserLicenses -join ", "

    # Extract the last successful sign-in timestamp
    $LastSignInRaw = $User.SignInActivity.LastSuccessfulSignInDateTime
    $LastSignIn = "No interactive sign-in recorded"
    if ($LastSignInRaw) {
        $ParsedSignIn = Parse-DateString $LastSignInRaw
        if ($null -ne $ParsedSignIn) {
            $LastSignIn = $ParsedSignIn.ToString("dd/MM/yyyy HH:mm")
        }
    }

    $Username = ($User.UserPrincipalName -split '@')[0].ToLower()
    $DisplayNameLower = $User.DisplayName.ToLower()
    $Verification = "-"
    if ($Username -like "*seriun*" -or $Username -eq "jp" -or $Username -like "jp.*" -or $Username -like "*.jp" -or $DisplayNameLower -like "*seriun*" -or $DisplayNameLower -match "\bjp\b") {
        $Verification = "Seriun/JP Account"
    }

    $Days = Get-DaysSince $LastSignIn
    $WastedCost = 0.00
    $MonthlySavings = 0.00
    if ($Days -gt 180) {
        $InactiveDays = $Days
        if ($Days -eq [double]::PositiveInfinity) {
            $InactiveDays = 365
        }
        $Months = $InactiveDays / 30
        $MonthlyPrice = Get-LicenseMonthlyPrice $LicenseString
        $WastedCost = $MonthlyPrice * $Months
        $MonthlySavings = $MonthlyPrice
    }
    $WastedCostText = "£{0:N2}" -f $WastedCost
    $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings

    # Downgrade Recommendation Engine
    $Recommendation = "-"
    if ($LicenseString) {
        $LicsArray = $LicenseString.Split(',') | ForEach-Object { $_.Trim() }
        $MonthlyPrice = Get-LicenseMonthlyPrice $LicenseString
        
        if ($Days -eq [double]::PositiveInfinity -or $Days -gt 180) {
            if ($MonthlyPrice -gt 0) {
                $Recommendation = "Reclaim: Remove all licenses (Save £{0:N2}/mo)" -f $MonthlyPrice
            }
        } else {
            # Check for redundant/overlapping licenses (e.g. having both a suite and standalone plans already included in it)
            $HasBusinessPremium = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Premium" }
            $HasM365E5 = $LicsArray | Where-Object { $_ -eq "Microsoft 365 E5" }
            $HasM365E3 = $LicsArray | Where-Object { $_ -eq "Microsoft 365 E3" }
            $HasO365E5 = $LicsArray | Where-Object { $_ -eq "Office 365 E5" }
            $HasO365E3 = $LicsArray | Where-Object { $_ -eq "Office 365 E3" }
            $HasBusinessStandard = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Standard" }
            $HasBusinessBasic = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Basic" }
            
            $RedundantPlans = [System.Collections.Generic.List[string]]::new()
            $RedundantCost = 0.00
            
            # 1. Exchange Online (Plan 1) redundant check
            $ExchangeP1Price = if ($SkuPrices.ContainsKey("Exchange Online (Plan 1)")) { $SkuPrices["Exchange Online (Plan 1)"] } else { 3.30 }
            if ($LicsArray -contains "Exchange Online (Plan 1)" -or $LicsArray -contains "EXCHANGESTANDARD") {
                if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3 -or $HasBusinessStandard -or $HasBusinessBasic) {
                    $RedundantPlans.Add("Exchange P1")
                    $RedundantCost += $ExchangeP1Price
                }
            }
            
            # 2. Exchange Online (Plan 2) redundant check (BP mailboxes upgrade to 100GB as of July 1, 2026!)
            $ExchangeP2Price = if ($SkuPrices.ContainsKey("Exchange Online (Plan 2)")) { $SkuPrices["Exchange Online (Plan 2)"] } else { 6.60 }
            if ($LicsArray -contains "Exchange Online (Plan 2)" -or $LicsArray -contains "EXCHANGEENTERPRISE") {
                if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                    $RedundantPlans.Add("Exchange P2")
                    $RedundantCost += $ExchangeP2Price
                }
            }
            
            # 3. Microsoft Entra ID P1 / AAD Premium redundant check
            $EntraP1Price = if ($SkuPrices.ContainsKey("Microsoft Entra ID P1")) { $SkuPrices["Microsoft Entra ID P1"] } else { 4.90 }
            if ($LicsArray -contains "Microsoft Entra ID P1" -or $LicsArray -contains "AAD_PREMIUM") {
                if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                    $RedundantPlans.Add("Entra ID P1")
                    $RedundantCost += $EntraP1Price
                }
            }
            
            # 4. Microsoft Intune redundant check
            $IntunePrice = if ($SkuPrices.ContainsKey("Microsoft Intune")) { $SkuPrices["Microsoft Intune"] } else { 6.60 }
            if ($LicsArray -contains "Microsoft Intune" -or $LicsArray -contains "INTUNE_A") {
                if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3) {
                    $RedundantPlans.Add("Intune")
                    $RedundantCost += $IntunePrice
                }
            }

            # 5. Power BI Pro redundant check
            $PowerBIPrice = if ($SkuPrices.ContainsKey("Power BI Pro")) { $SkuPrices["Power BI Pro"] } else { 8.20 }
            if ($LicsArray -contains "Power BI Pro" -or $LicsArray -contains "POWER_BI_PRO") {
                if ($HasM365E5 -or $HasO365E5) {
                    $RedundantPlans.Add("Power BI Pro")
                    $RedundantCost += $PowerBIPrice
                }
            }

            if ($RedundantPlans.Count -gt 0) {
                $PlansStr = $RedundantPlans -join " & "
                $Recommendation = "Redundant: Remove $PlansStr (Save £{0:N2}/mo)" -f $RedundantCost
                $MonthlySavings = $RedundantCost
                $WastedCost = $RedundantCost
                $WastedCostText = "£{0:N2}" -f $WastedCost
                $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings
            } else {
                # Check for Copilot inactive > 30 days
                $HasCopilot = $LicsArray | Where-Object { $_ -match "Copilot" }
                if ($HasCopilot -and $Days -gt 30) {
                    $Recommendation = "Downgrade: Remove Copilot (Save £24.70/mo)"
                } else {
                    # Check for E5 inactive > 90 days
                    $HasE5 = $LicsArray | Where-Object { $_ -match "E5" }
                    if ($HasE5 -and $Days -gt 90) {
                        $Recommendation = "Downgrade: E5 -> Business Premium (Save £30.00/mo)"
                    } else {
                        # Check for E3 inactive > 90 days
                        $HasE3 = $LicsArray | Where-Object { $_ -match "E3" }
                        if ($HasE3 -and $Days -gt 90) {
                            $Recommendation = "Downgrade: E3 -> Business Standard (Save £21.40/mo)"
                        }
                    }
                }
            }
        }
    }

    $Report.Add([PSCustomObject]@{
        "DisplayName"       = $User.DisplayName
        "UserPrincipalName" = $User.UserPrincipalName
        "AssignedLicenses"  = $LicenseString
        "LastSignInDate"    = $LastSignIn
        "WastedCost"        = $WastedCostText
        "MonthlySavings"    = $MonthlySavingsText
        "Recommendation"    = $Recommendation
        "Verification"      = $Verification
    })
}

# Output to GridView for quick inspection if the cmdlet is available
if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
    Write-Log "Displaying results in GridView..."
    $Report | Out-GridView -Title "Licensed Users Report - $TenantName"
} else {
    Write-Log "Out-GridView is not supported in this host environment. Skipping grid display." -Type Warning
}

# Export report to CSV log location
$HasSeriunOrJP = $Report | Where-Object { $_.Verification -eq "Seriun/JP Account" }
if ($HasSeriunOrJP) {
    Write-Log "Warning: Seriun or JP verification accounts were detected. Excluding them from the exported CSV file." -Type Warning
}
$FilteredReport = $Report | Where-Object { $_.Verification -ne "Seriun/JP Account" }
$FilteredReport | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding utf8
Write-Log "Report successfully exported to: $OutputFile" -Type Success

# Cleanly disconnect the Graph session
Write-Log "Disconnecting from Microsoft Graph..."
Disconnect-MgGraph | Out-Null
Write-Log "Execution finished." -ForegroundColor Green
Write-Log "--------------------------------------------------" -ForegroundColor DarkGray