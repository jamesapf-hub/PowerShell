<#
.SYNOPSIS
    Entra ID MFA SMS Deprecation Checker GUI & Color-Coded Excel Exporter
.DESCRIPTION
    Launches a modern WPF GUI to connect to Microsoft Entra ID via Microsoft Graph API,
    audit user MFA methods, check Legacy Per-User MFA State (Disabled/Enabled/Enforced) via Graph Beta /authentication/requirements,
    highlight users with SMS as Default or SMS as ONLY method, and export a color-coded XLSX report.
.EXAMPLE
    pwsh.exe -ExecutionPolicy Bypass -File .\Start-MFACheckerGUI.ps1
#>

# ==========================================
# 1. IN-MEMORY EXECUTION SAFETY GUARD
# ==========================================
if ($PSScriptRoot -match "^iex") {
    Write-Error "CRITICAL SAFETY ERROR: This packaged script relies on WPF GUI components and local modules. In-memory execution via 'iex' is disabled. Please download and extract the package folder, then run locally via:`n`npwsh.exe -ExecutionPolicy Bypass -File .\Start-MFACheckerGUI.ps1"
    exit 1
}

# Add required WPF & WinForms assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# Global State
$global:GraphConnected = $false
$global:TenantName = ""
$global:MfaAuditResults = [System.Collections.Generic.List[PSObject]]::new()

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
function Write-GuiLog ($msg) {
    if ($txtLog) {
        $txtLog.Dispatcher.Invoke([Action]{
            $txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $msg`n")
            $txtLog.ScrollToEnd()
        }, [System.Windows.Threading.DispatcherPriority]::Background)
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $msg"
    }
}

function Update-Progress ($val) {
    if ($pbStatus) {
        $pbStatus.Dispatcher.Invoke([Action]{
            $pbStatus.Value = $val
        }, [System.Windows.Threading.DispatcherPriority]::Background)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Refresh-KpiStats {
    if ($kpiTotal -and $dgAudit) {
        $total = $global:MfaAuditResults.Count
        $smsOnly = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SMS ONLY" }).Count
        $smsDefault = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SMS DEFAULT" }).Count
        $secure = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SECURE MFA" }).Count

        $kpiTotal.Text = $total.ToString()
        $kpiSmsOnly.Text = $smsOnly.ToString()
        $kpiSmsDefault.Text = $smsDefault.ToString()
        $kpiSecure.Text = $secure.ToString()

        $dgAudit.Dispatcher.Invoke([Action]{
            $dgAudit.ItemsSource = $null
            $dgAudit.ItemsSource = @($global:MfaAuditResults)
        }, [System.Windows.Threading.DispatcherPriority]::Background)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Format-FriendlyMfaMethod ($rawMethod, $registeredArray) {
    if ($rawMethod -and $rawMethod -notmatch "^(default|none|unknown|Not Configured)$") {
        switch -Regex ($rawMethod) {
            "push|microsoftAuthenticator"   { return "Microsoft Authenticator (Push)" }
            "sms|mobilePhone"               { return "SMS / Mobile Phone" }
            "voiceMobile|voiceOffice|phone" { return "Phone Call" }
            "oath|softwareOath"             { return "Software OATH (TOTP)" }
            "fido2|passkey"                 { return "FIDO2 Security Key / Passkey" }
            "windowsHello"                  { return "Windows Hello for Business" }
            "email"                         { return "Email" }
            default                         { return $rawMethod }
        }
    }

    if ($registeredArray) {
        $regStr = ($registeredArray -join " ")
        if ($regStr -match "fido2|passkey") {
            return "FIDO2 Passkey (System Default)"
        } elseif ($regStr -match "microsoftAuthenticator|push") {
            return "Microsoft Authenticator (System Default)"
        } elseif ($regStr -match "softwareOath|oath") {
            return "Software OATH TOTP (System Default)"
        } elseif ($regStr -match "mobilePhone|voiceMobile|voiceOffice|sms|phone") {
            return "SMS / Mobile Phone (System Default)"
        } elseif ($regStr -match "email") {
            return "Email (System Default)"
        }
    }

    return "Not Configured"
}

# ==========================================
# 3. MODULE LOADER & DEPENDENCY CHECK
# ==========================================
function Assert-RequiredModules {
    Write-GuiLog "Checking required PowerShell modules..."
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-GuiLog "ImportExcel module missing. Attempting automatic installation for CurrentUser..."
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-GuiLog "ImportExcel module successfully installed!"
        } catch {
            Write-GuiLog "WARNING: Failed to auto-install ImportExcel. Excel export will fall back to CSV if ImportExcel is unavailable."
        }
    } else {
        Write-GuiLog "ImportExcel module ready."
    }

    $graphAvailable = (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) -or (Get-Module -ListAvailable -Name Microsoft.Graph)
    if (-not $graphAvailable) {
        Write-GuiLog "WARNING: Microsoft.Graph modules not found. Ensure Microsoft.Graph is installed (`Install-Module Microsoft.Graph -Scope CurrentUser`)."
    } else {
        Write-GuiLog "Microsoft.Graph module ready."
    }
}

# ==========================================
# 4. GRAPH API & MFA AUDIT LOGIC
# ==========================================
function Connect-EntraIDGraph {
    Write-GuiLog "Initiating single-session Microsoft Graph authentication..."
    Write-GuiLog "Requesting permissions: User.Read.All, UserAuthenticationMethod.Read.All, Reports.Read.All, Directory.Read.All"

    try {
        # Check if already connected with a valid account session
        $existingContext = Get-MgContext
        if ($existingContext -and $existingContext.Account) {
            $global:GraphConnected = $true
            $global:TenantName = $existingContext.Account
            Write-GuiLog "Reusing active Graph session for account: $($existingContext.Account)"
            
            if ($txtConnStatus) {
                $txtConnStatus.Dispatcher.Invoke([Action]{
                    $txtConnStatus.Text = "Connected: $($existingContext.Account)"
                    $txtConnStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                })
            }
            return $true
        }

        # Clear stale tokens and request all required v1.0 and beta scopes in ONE upfront request
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        $scopes = @("User.Read.All", "UserAuthenticationMethod.Read.All", "Reports.Read.All", "Directory.Read.All")
        Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome -ErrorAction Stop

        $context = Get-MgContext
        if ($context) {
            $global:GraphConnected = $true
            $global:TenantName = $context.Account
            Write-GuiLog "Successfully connected to tenant: $($context.Account) (Tenant ID: $($context.TenantId))"
            
            if ($txtConnStatus) {
                $txtConnStatus.Dispatcher.Invoke([Action]{
                    $txtConnStatus.Text = "Connected: $($context.Account)"
                    $txtConnStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
                })
            }
            return $true
        }
    } catch {
        Write-GuiLog "ERROR: Graph Connection failed: $_"
        if ($txtConnStatus) {
            $txtConnStatus.Dispatcher.Invoke([Action]{
                $txtConnStatus.Text = "Connection Failed"
                $txtConnStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
            })
        }
        return $false
    }
}

function Invoke-MfaMethodAudit {
    param(
        [bool]$ActiveOnly = $true,
        [bool]$ExcludeGuests = $true
    )

    $global:MfaAuditResults.Clear()
    Refresh-KpiStats
    Write-GuiLog "Starting MFA Method Audit scan across tenant..."
    Update-Progress 10

    $regDetails = @()

    # Attempt 1: Direct REST API v1.0 /reports/authenticationMethods/userRegistrationDetails
    try {
        Write-GuiLog "Querying Graph REST API v1.0 /reports/authenticationMethods/userRegistrationDetails..."
        $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
        if ($response -and $response.value) {
            $regDetails = @($response.value)
            while ($response.'@odata.nextLink') {
                $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink' -ErrorAction SilentlyContinue
                if ($response.value) { $regDetails += $response.value }
            }
            Write-GuiLog "REST v1.0 returned $($regDetails.Count) user registration records."
        }
    } catch {
        Write-GuiLog "REST v1.0 query note: $_"
    }

    # Attempt 2: Direct REST API beta /reports/authenticationMethods/userRegistrationDetails
    if (-not $regDetails -or $regDetails.Count -eq 0) {
        try {
            Write-GuiLog "Querying Graph REST API beta /reports/authenticationMethods/userRegistrationDetails..."
            $uri = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
            if ($response -and $response.value) {
                $regDetails = @($response.value)
                while ($response.'@odata.nextLink') {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink' -ErrorAction SilentlyContinue
                    if ($response.value) { $regDetails += $response.value }
                }
                Write-GuiLog "REST beta returned $($regDetails.Count) user registration records."
            }
        } catch {
            Write-GuiLog "REST beta query note: $_"
        }
    }

    # Attempt 4: Fallback Directory Users query (/v1.0/users)
    if (-not $regDetails -or $regDetails.Count -eq 0) {
        try {
            Write-GuiLog "Fallback: Querying user directory accounts from /v1.0/users..."
            $usersUri = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,displayName,userType,accountEnabled&`$top=999"
            $uResp = Invoke-MgGraphRequest -Method GET -Uri $usersUri -ErrorAction SilentlyContinue
            if ($uResp -and $uResp.value) {
                $rawUsers = $uResp.value
                Write-GuiLog "Retrieved $($rawUsers.Count) user directory accounts. Querying authentication methods per user..."
                
                $idx = 0
                $regDetails = foreach ($u in $rawUsers) {
                    $idx++
                    if ($idx % 5 -eq 0 -or $idx -eq $rawUsers.Count) {
                        Update-Progress (10 + [int](($idx / $rawUsers.Count) * 40))
                    }
                    $mUri = "https://graph.microsoft.com/v1.0/users/$($u.id)/authentication/methods"
                    $mResp = Invoke-MgGraphRequest -Method GET -Uri $mUri -ErrorAction SilentlyContinue
                    $mTypes = @()
                    if ($mResp -and $mResp.value) {
                        $mTypes = $mResp.value | ForEach-Object {
                            $t = $_.'@odata.type' -replace '^#microsoft\.graph\.', ''
                            $t -replace 'AuthenticationMethod$', ''
                        }
                    }
                    
                    [PSCustomObject]@{
                        userPrincipalName   = $u.userPrincipalName
                        userDisplayName     = $u.displayName
                        userType            = $u.userType
                        accountEnabled      = $u.accountEnabled
                        isMfaRegistered     = ($mTypes.Count -gt 0 -and ($mTypes | Where-Object { $_ -ne "password" }).Count -gt 0)
                        isMfaCapable        = ($mTypes.Count -gt 0)
                        isSmsUser           = ($mTypes -contains "phone" -or $mTypes -contains "mobilePhone")
                        userPreferredMethod = if ($mTypes -contains "microsoftAuthenticator") { "push" } elseif ($mTypes -contains "phone" -or $mTypes -contains "mobilePhone") { "sms" } else { "default" }
                        methodsRegistered   = $mTypes
                    }
                }
            }
        } catch {
            Write-GuiLog "Directory users query note: $_"
        }
    }

    Update-Progress 40
    $totalCount = if ($regDetails) { $regDetails.Count } else { 0 }
    Write-GuiLog "Retrieved $totalCount total user records from Entra ID."

    if ($totalCount -eq 0) {
        Write-GuiLog "CRITICAL WARNING: 0 user records were returned. Ensure your admin account has consented to User.Read.All and UserAuthenticationMethod.Read.All."
        Update-Progress 0
        return $false
    }

    Update-Progress 50
    Write-GuiLog "Auditing per-user authentication requirements via Graph Beta (/beta/users/{id}/authentication/requirements)..."
    $processed = 0

    foreach ($user in $regDetails) {
        $processed++
        
        $upn = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { $user.userPrincipalName }
        $displayName = if ($user.UserDisplayName) { $user.UserDisplayName } else { $user.userDisplayName }
        $userType = if ($user.UserType) { $user.UserType } else { $user.userType }
        if (-not $userType) { $userType = "Member" }

        # Filter rules
        if ($ExcludeGuests -and ($userType -eq "Guest" -or $upn -like "*#EXT#*")) {
            continue
        }

        $isMfaRegistered = [bool]($user.IsMfaRegistered -or $user.isMfaRegistered)
        $isSmsUser = [bool]($user.IsSmsUser -or $user.isSmsUser)

        # Raw registered methods array
        $methods = @()
        if ($user.MethodsRegistered) {
            $methods = $user.MethodsRegistered
        } elseif ($user.methodsRegistered) {
            $methods = $user.methodsRegistered
        }

        # Human-friendly registered methods string
        $friendlyRegList = foreach ($m in $methods) {
            switch -Regex ($m) {
                "microsoftAuthenticator" { "Microsoft Authenticator (Push)" }
                "softwareOath"           { "Software OATH (TOTP)" }
                "mobilePhone"            { "SMS / Mobile Phone" }
                "voiceMobile"            { "Phone Call (Mobile)" }
                "voiceOffice"            { "Phone Call (Office)" }
                "fido2"                  { "FIDO2 Passkey" }
                "windowsHello"           { "Windows Hello for Business" }
                "email"                  { "Email" }
                "temporaryAccessPass"    { "Temporary Access Pass" }
                default                  { $m }
            }
        }
        $friendlyMethodsStr = ($friendlyRegList -join ", ")

        # Raw & Friendly Default MFA Method
        $rawPreferred = if ($user.UserPreferredMethod) { $user.UserPreferredMethod } else { $user.userPreferredMethod }
        $friendlyDefaultMfa = Format-FriendlyMfaMethod -rawMethod $rawPreferred -registeredArray $methods

        # Legacy Per-User MFA State via Graph Beta endpoint /beta/users/{encodedUpn}/authentication/requirements
        $legacyMfaState = "Disabled"
        if ($upn) {
            try {
                $encodedUpn = [Uri]::EscapeDataString($upn)
                $reqUri = "https://graph.microsoft.com/beta/users/${encodedUpn}/authentication/requirements"
                $reqResp = Invoke-MgGraphRequest -Method GET -Uri $reqUri -ErrorAction SilentlyContinue
                if ($reqResp -and $reqResp.perUserMfaState) {
                    $st = [string]$reqResp.perUserMfaState
                    if ($st -ieq "enforced") { $legacyMfaState = "Enforced" }
                    elseif ($st -ieq "enabled") { $legacyMfaState = "Enabled" }
                    elseif ($st -ieq "disabled") { $legacyMfaState = "Disabled" }
                    else { $legacyMfaState = (Get-Culture).TextInfo.ToTitleCase($st) }
                }
            } catch {
                # Fallback to Disabled if error
            }
        }

        # Determine SMS Default
        $isSmsDefault = ($friendlyDefaultMfa -like "*SMS*" -or $rawPreferred -in @("sms", "mobilePhone", "voiceMobile", "voiceOffice", "phone"))

        # Determine SMS Only
        $nonSmsMethods = $methods | Where-Object { $_ -notmatch "mobilePhone|voiceMobile|voiceOffice|sms|email|password|phone" }
        $hasSms = ($methods -match "mobilePhone|voiceMobile|voiceOffice|sms|phone") -or $isSmsUser
        $isSmsOnly = $hasSms -and ($nonSmsMethods.Count -eq 0)

        # Classify Risk Category
        $riskCategory = "SECURE MFA"
        $riskLevel = "Low"
        $recommendation = "Compliant. User is using strong MFA method."

        if (-not $isMfaRegistered -and $methods.Count -eq 0) {
            $riskCategory = "UNREGISTERED"
            $riskLevel = "Critical"
            $recommendation = "Require MFA registration via Conditional Access."
        } elseif ($isSmsOnly) {
            $riskCategory = "SMS ONLY"
            $riskLevel = "High"
            $recommendation = "CRITICAL ACTION: Register Microsoft Authenticator or Passkey before SMS deprecation."
        } elseif ($isSmsDefault) {
            $riskCategory = "SMS DEFAULT"
            $riskLevel = "Medium"
            $recommendation = "Prompt user to change default MFA method to Microsoft Authenticator Push."
        }

        # Columns: Display Name FIRST, UPN, Legacy MFA State
        $auditItem = [PSCustomObject]@{
            DisplayName          = $displayName
            UserPrincipalName    = $upn
            UserType             = $userType
            DefaultMfaMethod     = $friendlyDefaultMfa
            RegisteredMethods    = $friendlyMethodsStr
            LegacyMfaState       = $legacyMfaState
            RiskCategory         = $riskCategory
            RiskLevel            = $riskLevel
            Recommendation       = $recommendation
        }

        $global:MfaAuditResults.Add($auditItem)

        if ($processed % 15 -eq 0 -or $processed -eq $totalCount) {
            $pct = 50 + [int](($processed / $totalCount) * 45)
            Update-Progress $pct
            Refresh-KpiStats
        }
    }

    Update-Progress 100
    Refresh-KpiStats
    Write-GuiLog "MFA Method Audit Complete! Successfully displayed $($global:MfaAuditResults.Count) users in table."
    return $true
}

function Export-MfaAuditToExcel {
    param([string]$FilePath, [bool]$AutoOpen = $true)

    if ($global:MfaAuditResults.Count -eq 0) {
        Write-GuiLog "ERROR: No audit data available to export. Run scan first."
        return $false
    }

    Write-GuiLog "Exporting color-coded XLSX report to: $FilePath..."

    try {
        if (Test-Path $FilePath) {
            Remove-Item $FilePath -Force -ErrorAction SilentlyContinue
        }

        # Metrics for Summary tab
        $totalScanned = $global:MfaAuditResults.Count
        $smsOnlyCount = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SMS ONLY" }).Count
        $smsDefaultCount = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SMS DEFAULT" }).Count
        $secureCount = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "SECURE MFA" }).Count
        $unregCount = ($global:MfaAuditResults | Where-Object { $_.RiskCategory -eq "UNREGISTERED" }).Count
        $smsDependentPct = if ($totalScanned -gt 0) { [math]::Round((($smsOnlyCount + $smsDefaultCount) / $totalScanned) * 100, 1) } else { 0 }

        $summaryData = @(
            [PSCustomObject]@{ Metric = "Tenant Account"; Value = $global:TenantName },
            [PSCustomObject]@{ Metric = "Report Timestamp"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") },
            [PSCustomObject]@{ Metric = "Total Users Scanned"; Value = $totalScanned },
            [PSCustomObject]@{ Metric = "SMS ONLY Users (High Deprecation Risk)"; Value = $smsOnlyCount },
            [PSCustomObject]@{ Metric = "SMS Default Users (Medium Risk)"; Value = $smsDefaultCount },
            [PSCustomObject]@{ Metric = "Secure MFA Users (Compliant)"; Value = $secureCount },
            [PSCustomObject]@{ Metric = "Unregistered Users"; Value = $unregCount },
            [PSCustomObject]@{ Metric = "SMS Dependency Ratio"; Value = "$smsDependentPct %" }
        )

        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-GuiLog "Building Excel workbook with Executive Summary and color-coded Audit sheets..."
            
            # Export Summary Sheet
            $summaryData | Export-Excel -Path $FilePath -WorksheetName "Executive Summary" -AutoSize -TableStyle "Medium2" -FreezeTopRow

            # Export Audit Sheet
            $excelPackage = $global:MfaAuditResults | Export-Excel -Path $FilePath -WorksheetName "SMS Deprecation Audit" -AutoSize -TableStyle "Medium9" -FreezeTopRow -PassThru
            $ws = $excelPackage.Workbook.Worksheets["SMS Deprecation Audit"]
            $totalRows = $global:MfaAuditResults.Count

            # Define System.Drawing.Color objects for EPPlus compatibility
            $colorSmsOnlyFill   = [System.Drawing.ColorTranslator]::FromHtml("#FFC7CE") # Soft Red
            $colorSmsOnlyText   = [System.Drawing.ColorTranslator]::FromHtml("#9C0006") # Dark Red

            $colorSmsDefFill    = [System.Drawing.ColorTranslator]::FromHtml("#FFEB9C") # Soft Amber/Yellow
            $colorSmsDefText    = [System.Drawing.ColorTranslator]::FromHtml("#9C6500") # Dark Amber

            $colorSecureFill    = [System.Drawing.ColorTranslator]::FromHtml("#C6EFCE") # Soft Green
            $colorSecureText    = [System.Drawing.ColorTranslator]::FromHtml("#006100") # Dark Green

            $colorUnregFill     = [System.Drawing.ColorTranslator]::FromHtml("#E0E0E0") # Soft Gray
            $colorUnregText     = [System.Drawing.ColorTranslator]::FromHtml("#333333") # Dark Gray

            $colorHeaderFill    = [System.Drawing.ColorTranslator]::FromHtml("#1F4E78") # Navy Header
            $colorHeaderText    = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF") # White Text

            try {
                # Format Entire Header Row 1 across all columns
                $headerRow = $ws.Row(1)
                $headerRow.Style.Font.Bold = $true
                $headerRow.Style.Font.Color.SetColor($colorHeaderText)
                $headerRow.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $headerRow.Style.Fill.BackgroundColor.SetColor($colorHeaderFill)

                # Format Entire Data Rows across all columns ($ws.Row($rowNum))
                for ($r = 0; $r -lt $totalRows; $r++) {
                    $rowNum = $r + 2
                    $cat = $global:MfaAuditResults[$r].RiskCategory
                    $row = $ws.Row($rowNum)
                    $row.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid

                    if ($cat -eq "SMS ONLY") {
                        $row.Style.Fill.BackgroundColor.SetColor($colorSmsOnlyFill)
                        $row.Style.Font.Color.SetColor($colorSmsOnlyText)
                        $row.Style.Font.Bold = $true
                    } elseif ($cat -eq "SMS DEFAULT") {
                        $row.Style.Fill.BackgroundColor.SetColor($colorSmsDefFill)
                        $row.Style.Font.Color.SetColor($colorSmsDefText)
                    } elseif ($cat -eq "SECURE MFA") {
                        $row.Style.Fill.BackgroundColor.SetColor($colorSecureFill)
                        $row.Style.Font.Color.SetColor($colorSecureText)
                    } elseif ($cat -eq "UNREGISTERED") {
                        $row.Style.Fill.BackgroundColor.SetColor($colorUnregFill)
                        $row.Style.Font.Color.SetColor($colorUnregText)
                    }
                }
            } catch {
                Write-GuiLog "Full row styling note: $_"
            }

            Close-ExcelPackage $excelPackage
            Write-GuiLog "Color-coded XLSX export generated successfully with full row highlighting!"
        } else {
            $csvPath = $FilePath -replace "\.xlsx$", ".csv"
            $global:MfaAuditResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-GuiLog "ImportExcel unavailable. Exported fallback CSV file to: $csvPath"
            $FilePath = $csvPath
        }

        if ($AutoOpen -and (Test-Path $FilePath)) {
            Invoke-Item $FilePath
        }

        return $true
    } catch {
        Write-GuiLog "ERROR during Excel export: $_"
        return $false
    }
}

# ==========================================
# 5. WPF XAML GUI CONSTRUCTION
# ==========================================
$inputXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Entra ID MFA SMS Deprecation Checker" Height="780" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#1E1E2E" Foreground="#CDD6F4"
        FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#181825"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderBrush" Value="#45475A"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="4"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="Margin" Value="6,4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Title & Connection Bar -->
            <RowDefinition Height="Auto"/> <!-- KPI Cards -->
            <RowDefinition Height="Auto"/> <!-- Controls Bar -->
            <RowDefinition Height="*"/>    <!-- DataGrid Preview -->
            <RowDefinition Height="Auto"/> <!-- Export Bar -->
            <RowDefinition Height="120"/>  <!-- Console Log -->
        </Grid.RowDefinitions>

        <!-- HEADER BAR -->
        <Border Grid.Row="0" Background="#252538" CornerRadius="8" Padding="12" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Orientation="Vertical">
                    <TextBlock Text="Entra ID MFA SMS Deprecation Checker" FontSize="20" FontWeight="Bold" Foreground="#89B4FA"/>
                    <TextBlock Text="Audit user authentication methods, identify SMS default &amp; SMS-only users, and export color-coded XLSX" FontSize="12" Foreground="#A6ADC8" Margin="0,2,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="txtConnStatus" Text="Not Connected" Foreground="Tomato" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <Button Name="btnConnect" Content="Connect to Entra ID" Background="#89B4FA" Foreground="#11111B" Padding="16,8"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- KPI CARDS -->
        <Grid Grid.Row="1" Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Total Users -->
            <Border Grid.Column="0" Background="#252538" CornerRadius="6" Padding="12" Margin="0,0,6,0">
                <StackPanel>
                    <TextBlock Text="Total Users Scanned" FontSize="11" Foreground="#A6ADC8"/>
                    <TextBlock Name="kpiTotal" Text="0" FontSize="24" FontWeight="Bold" Foreground="#89B4FA" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>

            <!-- SMS ONLY -->
            <Border Grid.Column="1" Background="#252538" CornerRadius="6" Padding="12" Margin="3,0,3,0">
                <StackPanel>
                    <TextBlock Text="SMS ONLY (High Risk)" FontSize="11" Foreground="#F38BA8"/>
                    <TextBlock Name="kpiSmsOnly" Text="0" FontSize="24" FontWeight="Bold" Foreground="#F38BA8" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>

            <!-- SMS DEFAULT -->
            <Border Grid.Column="2" Background="#252538" CornerRadius="6" Padding="12" Margin="3,0,3,0">
                <StackPanel>
                    <TextBlock Text="SMS Default (Medium Risk)" FontSize="11" Foreground="#F9E2AF"/>
                    <TextBlock Name="kpiSmsDefault" Text="0" FontSize="24" FontWeight="Bold" Foreground="#F9E2AF" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>

            <!-- SECURE MFA -->
            <Border Grid.Column="3" Background="#252538" CornerRadius="6" Padding="12" Margin="6,0,0,0">
                <StackPanel>
                    <TextBlock Text="Secure MFA (Compliant)" FontSize="11" Foreground="#A6E3A1"/>
                    <TextBlock Name="kpiSecure" Text="0" FontSize="24" FontWeight="Bold" Foreground="#A6E3A1" Margin="0,4,0,0"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- CONTROLS BAR -->
        <Border Grid.Row="2" Background="#252538" CornerRadius="6" Padding="8" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <CheckBox Name="chkActiveOnly" Content="Active Users Only" IsChecked="True" Grid.Column="0"/>
                <CheckBox Name="chkExcludeGuests" Content="Exclude Guest Users" IsChecked="True" Grid.Column="1"/>

                <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
                    <TextBlock Text="Search:" VerticalAlignment="Center" Foreground="#A6ADC8" Margin="0,0,4,0"/>
                    <TextBox Name="txtSearch" Width="200"/>
                </StackPanel>

                <Button Name="btnScan" Content="Scan MFA Methods" Grid.Column="3" Background="#A6E3A1" Foreground="#11111B" Padding="16,6"/>
            </Grid>
        </Border>

        <!-- DATAGRID PREVIEW WITH FULL ROW COLOR CODING -->
        <Border Grid.Row="3" Background="#181825" CornerRadius="6" Margin="0,0,0,8">
            <DataGrid Name="dgAudit" AutoGenerateColumns="False" IsReadOnly="True"
                      Background="#181825" Foreground="#CDD6F4" RowBackground="#1E1E2E"
                      AlternatingRowBackground="#252538" GridLinesVisibility="Horizontal"
                      HorizontalGridLinesBrush="#313244" HeadersVisibility="Column">
                <DataGrid.Resources>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="#313244"/>
                        <Setter Property="Foreground" Value="#CDD6F4"/>
                        <Setter Property="Padding" Value="8,6"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                    </Style>
                </DataGrid.Resources>

                <!-- Full Row Highlighting Triggers for GUI DataGrid -->
                <DataGrid.RowStyle>
                    <Style TargetType="DataGridRow">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding RiskCategory}" Value="SMS ONLY">
                                <Setter Property="Background" Value="#FFC7CE"/>
                                <Setter Property="Foreground" Value="#9C0006"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding RiskCategory}" Value="SMS DEFAULT">
                                <Setter Property="Background" Value="#FFEB9C"/>
                                <Setter Property="Foreground" Value="#9C6500"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding RiskCategory}" Value="SECURE MFA">
                                <Setter Property="Background" Value="#C6EFCE"/>
                                <Setter Property="Foreground" Value="#006100"/>
                            </DataTrigger>
                            <DataTrigger Binding="{Binding RiskCategory}" Value="UNREGISTERED">
                                <Setter Property="Background" Value="#E0E0E0"/>
                                <Setter Property="Foreground" Value="#333333"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </DataGrid.RowStyle>

                <DataGrid.Columns>
                    <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="1.8*"/>
                    <DataGridTextColumn Header="User Principal Name" Binding="{Binding UserPrincipalName}" Width="2.2*"/>
                    <DataGridTextColumn Header="User Type" Binding="{Binding UserType}" Width="1*"/>
                    <DataGridTextColumn Header="Default MFA Method" Binding="{Binding DefaultMfaMethod}" Width="1.8*"/>
                    <DataGridTextColumn Header="Registered Methods" Binding="{Binding RegisteredMethods}" Width="2*"/>
                    <DataGridTextColumn Header="Legacy MFA State" Binding="{Binding LegacyMfaState}" Width="1.2*"/>
                    <DataGridTextColumn Header="Risk Category" Binding="{Binding RiskCategory}" Width="1.2*"/>
                    <DataGridTextColumn Header="Recommendation" Binding="{Binding Recommendation}" Width="2.5*"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>

        <!-- EXPORT BAR -->
        <Border Grid.Row="4" Background="#252538" CornerRadius="6" Padding="8" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Export Path:" VerticalAlignment="Center" Foreground="#A6ADC8" Margin="4,0,8,0"/>
                <TextBox Name="txtExportPath" Grid.Column="1"/>
                <Button Name="btnBrowse" Content="Browse..." Grid.Column="2"/>
                <Button Name="btnExport" Content="Export Color-Coded XLSX" Grid.Column="3" Background="#FAB387" Foreground="#11111B" Padding="16,6"/>
            </Grid>
        </Border>

        <!-- CONSOLE LOG & PROGRESS -->
        <Grid Grid.Row="5">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <ProgressBar Name="pbStatus" Height="6" Background="#313244" Foreground="#89B4FA" Margin="0,0,0,4"/>
            <TextBox Name="txtLog" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                     Background="#11111B" Foreground="#A6E3A1" FontFamily="Consolas" FontSize="11"/>
        </Grid>
    </Grid>
</Window>
"@

# Read XAML
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($inputXaml))
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Map UI Controls
$txtConnStatus  = $window.FindName("txtConnStatus")
$btnConnect     = $window.FindName("btnConnect")
$kpiTotal       = $window.FindName("kpiTotal")
$kpiSmsOnly     = $window.FindName("kpiSmsOnly")
$kpiSmsDefault  = $window.FindName("kpiSmsDefault")
$kpiSecure      = $window.FindName("kpiSecure")
$chkActiveOnly  = $window.FindName("chkActiveOnly")
$chkExcludeGuests = $window.FindName("chkExcludeGuests")
$txtSearch      = $window.FindName("txtSearch")
$btnScan        = $window.FindName("btnScan")
$dgAudit        = $window.FindName("dgAudit")
$txtExportPath  = $window.FindName("txtExportPath")
$btnBrowse      = $window.FindName("btnBrowse")
$btnExport      = $window.FindName("btnExport")
$pbStatus       = $window.FindName("pbStatus")
$txtLog         = $window.FindName("txtLog")

# Default Export Path
$defaultPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "MFA_SMS_Deprecation_Report_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx")
$txtExportPath.Text = $defaultPath

# Initial Log
$txtLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] Entra ID MFA SMS Deprecation Checker initialized.`n")
Assert-RequiredModules

# Helper to run scan process reliably with live UI updates
function Start-ScanProcess {
    if (-not $global:GraphConnected) {
        $connOk = Connect-EntraIDGraph
        if (-not $connOk) { return }
    }

    $btnScan.IsEnabled = $false
    $btnExport.IsEnabled = $false
    
    $act = [bool]$chkActiveOnly.IsChecked
    $gst = [bool]$chkExcludeGuests.IsChecked

    try {
        Invoke-MfaMethodAudit -ActiveOnly $act -ExcludeGuests $gst
    } finally {
        $btnScan.IsEnabled = $true
        $btnExport.IsEnabled = $true
    }
}

# EVENT HANDLERS

# 1. Connect Button (Connects and automatically scans/loads user list)
$btnConnect.Add_Click({
    $btnConnect.IsEnabled = $false
    try {
        Start-ScanProcess
    } finally {
        $btnConnect.IsEnabled = $true
    }
})

# 2. Scan Button (Connects if missing, then executes audit)
$btnScan.Add_Click({
    Start-ScanProcess
})

# 3. Browse Button
$btnBrowse.Add_Click({
    $sfd = [System.Windows.Forms.SaveFileDialog]::new()
    $sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $sfd.FileName = [System.IO.Path]::GetFileName($txtExportPath.Text)
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExportPath.Text = $sfd.FileName
    }
})

# 4. Export Button
$btnExport.Add_Click({
    $filePath = $txtExportPath.Text
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        [System.Windows.MessageBox]::Show("Please specify a valid Excel export path.", "Invalid Path", "OK", "Warning")
        return
    }

    $btnExport.IsEnabled = $false
    Export-MfaAuditToExcel -FilePath $filePath -AutoOpen $true
    $btnExport.IsEnabled = $true
})

# 5. Search Box Live Filter
$txtSearch.Add_TextChanged({
    $q = $txtSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($q)) {
        $dgAudit.ItemsSource = @($global:MfaAuditResults)
    } else {
        $filtered = $global:MfaAuditResults | Where-Object {
            $_.DisplayName -like "*$q*" -or $_.UserPrincipalName -like "*$q*" -or $_.DefaultMfaMethod -like "*$q*" -or $_.LegacyMfaState -like "*$q*" -or $_.RiskCategory -like "*$q*"
        }
        $dgAudit.ItemsSource = @($filtered)
    }
})

# Show GUI Window
[void]$window.ShowDialog()
