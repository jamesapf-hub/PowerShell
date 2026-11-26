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
# 1. IN-MEMORY EXECUTION SAFETY GUARD & STA CHECK
# ==========================================
if ($PSScriptRoot -match "^iex") {
    Write-Error "CRITICAL SAFETY ERROR: This packaged script relies on WPF GUI components and local modules. In-memory execution via 'iex' is disabled. Please download and extract the package folder, then run locally via:`n`npowershell.exe -ExecutionPolicy Bypass -File .\Start-MFACheckerGUI.ps1"
    exit 1
}

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
    exit
}

# Add required WPF & WinForms assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# Painted Windows 11 Chrome Helper (Frame / DwmSetWindowAttribute Pattern)
if (-not ('Win11ChromeHelper' -as [type])) {
    $chromeCSharp = @'
using System;
using System.Runtime.InteropServices;

public class Win11ChromeHelper
{
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
    private const int DWMWA_BORDER_COLOR = 34;
    private const int DWMWA_CAPTION_COLOR = 35;
    private const int DWMWA_TEXT_COLOR = 36;
    private const int DWMWCP_ROUND = 2;

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref uint attrValue, int attrSize);

    public static void ApplyWindowChrome(IntPtr hwnd, bool isDark)
    {
        if (hwnd == IntPtr.Zero) return;

        try
        {
            int darkVal = isDark ? 1 : 0;
            DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkVal, sizeof(int));

            int cornerVal = DWMWCP_ROUND;
            DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref cornerVal, sizeof(int));

            uint captionBgr = isDark ? 0x0026160Fu : 0x00FFFFFFu; // #0F1626 dark / #FFFFFF light
            uint textBgr    = isDark ? 0x00FCFAF8u : 0x002A170Fu; // #F8FAFC dark / #0F172A light
            uint borderBgr  = isDark ? 0x003B291Eu : 0x00E1D5CBu; // #1E293B dark / #CBD5E1 light

            DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ref captionBgr, sizeof(uint));
            DwmSetWindowAttribute(hwnd, DWMWA_TEXT_COLOR, ref textBgr, sizeof(uint));
            DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR, ref borderBgr, sizeof(uint));
        }
        catch
        {
            // Ignore gracefully on Windows 10 hosts lacking Win11 DWM attributes
        }
    }
}
'@
    try { Add-Type -TypeDefinition $chromeCSharp -Language CSharp -ErrorAction Stop } catch {}
}

function Get-SystemIsDarkTheme {
    [CmdletBinding()]
    param()
    try {
        $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        if (Test-Path -Path $regPath) {
            $val = Get-ItemPropertyValue -Path $regPath -Name 'AppsUseLightTheme' -ErrorAction Stop
            return ([int]$val -eq 0)
        }
    } catch {}
    return $true
}

function Set-StandardGuiTheme {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,

        [Parameter(Mandatory = $true)]
        [bool]$IsDark,

        [System.Windows.Controls.TextBlock]$ThemeLabel = $null,
        [bool]$IsSystemAuto = $true
    )

    $conv = New-Object System.Windows.Media.BrushConverter
    if ($IsDark) {
        $canvasBrush = New-Object System.Windows.Media.RadialGradientBrush
        $canvasBrush.Center = New-Object System.Windows.Point(0.38, 0.0)
        $canvasBrush.GradientOrigin = New-Object System.Windows.Point(0.38, 0.0)
        $canvasBrush.RadiusX = 0.95
        $canvasBrush.RadiusY = 0.9
        $canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#0F1F3D')), 0.0))
        $canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#091122')), 0.52))
        $canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#05070C')), 1.0))

        $map = @{
            BrushChromeBar     = '#CC070D19'
            BrushHairline      = '#1CFFFFFF'
            BrushCardBg        = '#B30D192E'
            BrushCardBorder    = '#26354D'
            BrushInputBg       = '#C7070D1A'
            BrushLogBg         = '#E6050A14'
            BrushTextPrimary   = '#F8FAFC'
            BrushTextSecondary = '#CBD5E1'
            BrushTextMuted     = '#94A3B8'
            BrushAccent        = '#2DD4BF'
            BrushAccentHover   = '#5EEAD4'
            BrushAccentPressed = '#14B8A6'
            BrushAccentFg      = '#042F2E'
            BrushAccentTint    = '#242DD4BF'
            BrushAccentBorder  = '#6B2DD4BF'
            BrushSecBtnBg      = '#14FFFFFF'
            BrushSecBtnHover   = '#24FFFFFF'
            BrushSecBtnPressed = '#0DFFFFFF'
            BrushScrollThumb   = '#26354D'
            BrushScrollTrack   = '#050A14'
            BrushDisabledBg    = '#1A2B44'
            BrushOk            = '#34D399'
            BrushWarn          = '#FBBF24'
            BrushErr           = '#F87171'
            BrushGridHeaderBg  = '#0B1528'
            BrushGridRowBg     = '#091122'
            BrushGridAltRowBg  = '#0C172B'
            BrushGridSelectedBg = '#1E3A5F'
            BrushGridLine      = '#1A2B44'
            BrushDangerBg      = '#26EF4444'
            BrushDangerBorder  = '#66EF4444'
            BrushDangerFg      = '#FCA5A5'
            BrushRowSmsOnlyBg  = '#2BEF4444'
            BrushRowSmsOnlyFg  = '#FCA5A5'
            BrushRowSmsDefBg   = '#24FBBF24'
            BrushRowSmsDefFg   = '#FDE68A'
            BrushRowSecureBg   = '#1E34D399'
            BrushRowSecureFg   = '#6EE7B7'
            BrushRowUnregBg    = '#1A94A3B8'
            BrushRowUnregFg    = '#CBD5E1'
        }
    } else {
        $canvasBrush = $conv.ConvertFromString('#F8F9FF')
        $map = @{
            BrushChromeBar     = '#F2FFFFFF'
            BrushHairline      = '#190F172A'
            BrushCardBg        = '#FFFFFFFF'
            BrushCardBorder    = '#CBD5E1'
            BrushInputBg       = '#F8FAFC'
            BrushLogBg         = '#F1F5F9'
            BrushTextPrimary   = '#0F172A'
            BrushTextSecondary = '#334155'
            BrushTextMuted     = '#64748B'
            BrushAccent        = '#2563EB'
            BrushAccentHover   = '#1D4ED8'
            BrushAccentPressed = '#1E40AF'
            BrushAccentFg      = '#FFFFFF'
            BrushAccentTint    = '#DBEAFE'
            BrushAccentBorder  = '#592563EB'
            BrushSecBtnBg      = '#FFFFFF'
            BrushSecBtnHover   = '#F1F5F9'
            BrushSecBtnPressed = '#E2E8F0'
            BrushScrollThumb   = '#CBD5E1'
            BrushScrollTrack   = '#F1F5F9'
            BrushDisabledBg    = '#E2E8F0'
            BrushOk            = '#059669'
            BrushWarn          = '#B45309'
            BrushErr           = '#DC2626'
            BrushGridHeaderBg  = '#F1F5F9'
            BrushGridRowBg     = '#FFFFFF'
            BrushGridAltRowBg  = '#F8FAFC'
            BrushGridSelectedBg = '#DBEAFE'
            BrushGridLine      = '#E2E8F0'
            BrushDangerBg      = '#FEE2E2'
            BrushDangerBorder  = '#FCA5A5'
            BrushDangerFg      = '#B91C1C'
            BrushRowSmsOnlyBg  = '#FEE2E2'
            BrushRowSmsOnlyFg  = '#991B1B'
            BrushRowSmsDefBg   = '#FEF3C7'
            BrushRowSmsDefFg   = '#92400E'
            BrushRowSecureBg   = '#D1FAE5'
            BrushRowSecureFg   = '#065F46'
            BrushRowUnregBg    = '#F1F5F9'
            BrushRowUnregFg    = '#334155'
        }
    }

    if ($canvasBrush.CanFreeze) { $canvasBrush.Freeze() }
    $Window.Background = $canvasBrush
    $Window.Resources['BrushCanvas'] = $canvasBrush
    foreach ($key in $map.Keys) {
        $b = $conv.ConvertFromString($map[$key])
        if ($b.CanFreeze) { $b.Freeze() }
        $Window.Resources[$key] = $b
    }

    if ($null -ne $ThemeLabel) {
        $modeTag = if ($IsSystemAuto) { 'System' } else { 'Manual' }
        $toneTag = if ($IsDark) { 'Dark' } else { 'Light' }
        $ThemeLabel.Text = "Theme: $modeTag ($toneTag)"
    }

    try {
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        if ($helper.Handle -ne [IntPtr]::Zero -and ('Win11ChromeHelper' -as [type])) {
            [Win11ChromeHelper]::ApplyWindowChrome($helper.Handle, $IsDark)
        }
    } catch {}
}

function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,
        [double]$MinWidth = 860,
        [double]$MinHeight = 640
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $fontDir = $null

    if ($PSScriptRoot) {
        $probe = $PSScriptRoot
        for ($i = 0; $i -le 5; $i++) {
            if ([string]::IsNullOrWhiteSpace($probe)) { break }
            $candidateDir = Join-Path -Path $probe -ChildPath 'assets\fonts'
            $candidateRegular = Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf'
            if (Test-Path -Path $candidateRegular -PathType Leaf) {
                $fontDir = $candidateDir
                break
            }
            $probe = Split-Path -Path $probe -Parent
        }
    }

    if (-not $fontDir) {
        $cacheFontDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        $cachedRegular = Join-Path -Path $cacheFontDir -ChildPath 'Figtree-Regular.ttf'
        if (-not (Test-Path -Path $cachedRegular -PathType Leaf)) {
            try {
                if (-not (Test-Path -Path $cacheFontDir)) {
                    New-Item -Path $cacheFontDir -ItemType Directory -Force | Out-Null
                }
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    $baseUrl = 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/'
                    foreach ($f in $fontFiles) {
                        $dest = Join-Path -Path $cacheFontDir -ChildPath $f
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            Invoke-WebRequest -Uri ($baseUrl + $f) -OutFile $dest -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                        }
                    }
                } finally {
                    $ProgressPreference = $prevProgress
                }
            } catch {
                Write-Verbose "Figtree font download skipped (offline fallback): $($_.Exception.Message)"
            }
        }
        if (Test-Path -Path $cachedRegular -PathType Leaf) {
            $fontDir = $cacheFontDir
        }
    }

    if ($fontDir) {
        try {
            $resolvedFontDir = (Resolve-Path -Path $fontDir -ErrorAction Stop).Path
            $fontUri = ([System.Uri]::new((Join-Path -Path $resolvedFontDir -ChildPath ''))).AbsoluteUri
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
        } catch {
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
        }
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Window.WindowStartupLocation = 'Manual'
    $Window.Width  = $targetWidth
    $Window.Height = $targetHeight
    $Window.MinWidth  = [Math]::Min($MinWidth, $targetWidth)
    $Window.MinHeight = [Math]::Min($MinHeight, $targetHeight)
    $Window.Left = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $Window.Top  = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Window.ResizeMode = 'CanResize'
}

# Global State
$global:GraphConnected = $false
$global:TenantName = ""
$global:MfaAuditResults = [System.Collections.Generic.List[PSObject]]::new()

# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
function Write-GuiLog ($msg) {
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    Write-Host "[$timestamp] $msg"
    if ($txtLog) {
        try {
            $txtLog.Dispatcher.Invoke([Action]{
                $txtLog.AppendText("[$timestamp] $msg`n")
                $txtLog.ScrollToEnd()
            }, [System.Windows.Threading.DispatcherPriority]::Background)
            [System.Windows.Forms.Application]::DoEvents()
        } catch {}
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
    Write-GuiLog "Checking required PowerShell modules and dependencies..."
    
    # Enable TLS 1.2 for PSGallery access across all PowerShell versions
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch {}

    # Check NuGet PackageProvider (required on clean Windows / PowerShell setups to install modules)
    try {
        $nugetVersion = try {
            Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction Stop | 
                Where-Object { $_.Version -ge [version]"2.8.5.201" }
        } catch {
            $null
        }
        if (-not $nugetVersion) {
            Write-GuiLog "NuGet provider missing or outdated. Installing NuGet provider for CurrentUser..."
            try {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            } catch {}
        }
    } catch {
        Write-GuiLog "NuGet provider note: $_"
    }

    # Ensure PSGallery repository is trusted to prevent interactive blocking prompts
    try {
        $psg = Get-PSRepository -Name "PSGallery" -ErrorAction Stop
        if ($psg -and $psg.InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction Stop
        }
    } catch {}

    # 1. Check ImportExcel
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-GuiLog "ImportExcel module missing. Attempting automatic installation from PSGallery..."
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            try {
                Import-Module ImportExcel -ErrorAction Stop
            } catch {}
            Write-GuiLog "ImportExcel module successfully installed and imported!"
        } catch {
            Write-GuiLog "WARNING: Failed to auto-install ImportExcel ($_). Excel export will fall back to CSV."
        }
    } else {
        try {
            Import-Module ImportExcel -ErrorAction Stop
        } catch {}
        Write-GuiLog "ImportExcel module ready."
    }

    # 2. Check Microsoft.Graph.Authentication (or full Microsoft.Graph)
    $graphAvailable = (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) -or (Get-Module -ListAvailable -Name Microsoft.Graph)
    if (-not $graphAvailable) {
        Write-GuiLog "Microsoft Graph modules not found. Installing Microsoft.Graph.Authentication for CurrentUser..."
        Write-GuiLog "(This is a fast, lightweight package containing Connect-MgGraph and Invoke-MgGraphRequest)"
        try {
            Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Write-GuiLog "Microsoft.Graph.Authentication module successfully installed and imported!"
        } catch {
            Write-GuiLog "ERROR: Failed to auto-install Microsoft.Graph.Authentication: $_"
            Write-GuiLog "To install manually, run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
            return $false
        }
    } else {
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        } catch {}
        Write-GuiLog "Microsoft Graph module ready."
    }

    return $true
}

# ==========================================
# 4. GRAPH API & MFA AUDIT LOGIC
# ==========================================
function Connect-EntraIDGraph {
    Write-GuiLog "Initiating single-session Microsoft Graph authentication..."
    Write-GuiLog "Requesting permissions: User.Read.All, UserAuthenticationMethod.Read.All, Reports.Read.All, AuditLog.Read.All, Directory.Read.All"

    # Safety check: Ensure Connect-MgGraph is available before invoking
    if (-not (try { Get-Command Connect-MgGraph -ErrorAction Stop } catch { $null })) {
        Write-GuiLog "Connect-MgGraph cmdlet not loaded in session. Verifying modules..."
        $modReady = Assert-RequiredModules
        if (-not $modReady -or -not (try { Get-Command Connect-MgGraph -ErrorAction Stop } catch { $null })) {
            Write-GuiLog "ERROR: Connect-MgGraph is still unavailable. Please run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
            if ($txtConnStatus) {
                $txtConnStatus.Dispatcher.Invoke([Action]{
                    $txtConnStatus.Text = "Missing Modules"
                    $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushErr')
                })
            }
            return $false
        }
    }

    try {
        # Force a clean state upfront by disconnecting any existing session
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
        } catch {}
        $global:GraphConnected = $false
        $global:TenantName = ""

        # Request all required v1.0 and beta scopes targeting 'organizations' authority for automatic tenant resolution
        $scopes = @("User.Read.All", "UserAuthenticationMethod.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Directory.Read.All")
        Connect-MgGraph -Scopes $scopes -TenantId "organizations" -ContextScope Process -NoWelcome -ErrorAction Stop

        $context = Get-MgContext
        if ($context) {
            $global:GraphConnected = $true
            $global:TenantName = $context.Account
            Write-GuiLog "Successfully connected to tenant: $($context.Account) (Tenant ID: $($context.TenantId))"
            
            if ($txtConnStatus) {
                $txtConnStatus.Dispatcher.Invoke([Action]{
                    $txtConnStatus.Text = "Connected: $($context.Account)"
                    $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushOk')
                })
            }
            return $true
        }
    } catch {
        Write-GuiLog "ERROR: Graph Connection failed: $_"
        if ($txtConnStatus) {
            $txtConnStatus.Dispatcher.Invoke([Action]{
                $txtConnStatus.Text = "Connection Failed"
                $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushErr')
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
        $uri = "v1.0/reports/authenticationMethods/userRegistrationDetails"
        $response = try { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop } catch { $null }
        if ($response -and $response.value) {
            $regDetails = @($response.value)
            while ($response.'@odata.nextLink') {
                $nextUri = $response.'@odata.nextLink' -replace '^https://graph\.microsoft\.com/', ''
                $response = try { Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop } catch { $null }
                if ($response.value) { $regDetails += $response.value }
            }
            Write-GuiLog "REST v1.0 returned $($regDetails.Count) user registration records."
        }
    } catch {
        Write-GuiLog "REST v1.0 query note: $_"
        if ("$_" -like "*AuditLog.Read.All*") {
            Write-GuiLog "NOTICE: AuditLog.Read.All permission is required for fast bulk reporting. Reconnect to grant consent."
        }
    }


    # Attempt 4: Fallback Directory Users query (/v1.0/users)
    if (-not $regDetails -or $regDetails.Count -eq 0) {
        try {
            Write-GuiLog "Fallback: Querying user directory accounts from /v1.0/users..."
            $usersUri = "v1.0/users?`$select=id,userPrincipalName,displayName,userType,accountEnabled&`$top=999"
            $uResp = try { Invoke-MgGraphRequest -Method GET -Uri $usersUri -ErrorAction Stop } catch { $null }
            if ($uResp -and $uResp.value) {
                $rawUsers = $uResp.value
                Write-GuiLog "Retrieved $($rawUsers.Count) user directory accounts. Querying authentication methods per user..."
                
                $idx = 0
                $regDetails = foreach ($u in $rawUsers) {
                    $idx++
                    if ($idx % 5 -eq 0 -or $idx -eq $rawUsers.Count) {
                        Update-Progress (10 + [int](($idx / $rawUsers.Count) * 40))
                    }
                    $mUri = "v1.0/users/$($u.id)/authentication/methods"
                    $mResp = try { Invoke-MgGraphRequest -Method GET -Uri $mUri -ErrorAction Stop } catch { $null }
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
                $reqUri = "beta/users/${encodedUpn}/authentication/requirements"
                $reqResp = Invoke-MgGraphRequest -Method GET -Uri $reqUri -ErrorAction Stop
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
            try {
                Remove-Item $FilePath -Force -ErrorAction Stop
            } catch {}
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
            [PSCustomObject]@{ Metric = "Report Timestamp"; Value = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss") },
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
        Title="Entra ID MFA SMS Deprecation Checker" Height="780" Width="1100" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual" Background="#091122" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13">
    <Window.Resources>
        <SolidColorBrush x:Key="BrushCanvas" Color="#091122"/>
        <SolidColorBrush x:Key="BrushChromeBar" Color="#CC070D19"/>
        <SolidColorBrush x:Key="BrushHairline" Color="#1CFFFFFF"/>
        <SolidColorBrush x:Key="BrushCardBg" Color="#B30D192E"/>
        <SolidColorBrush x:Key="BrushCardBorder" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushInputBg" Color="#C7070D1A"/>
        <SolidColorBrush x:Key="BrushLogBg" Color="#E6050A14"/>
        <SolidColorBrush x:Key="BrushTextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BrushTextSecondary" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="BrushTextMuted" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BrushAccent" Color="#2DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentHover" Color="#5EEAD4"/>
        <SolidColorBrush x:Key="BrushAccentPressed" Color="#14B8A6"/>
        <SolidColorBrush x:Key="BrushAccentFg" Color="#042F2E"/>
        <SolidColorBrush x:Key="BrushAccentTint" Color="#242DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentBorder" Color="#6B2DD4BF"/>
        <SolidColorBrush x:Key="BrushSecBtnBg" Color="#14FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnHover" Color="#24FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnPressed" Color="#0DFFFFFF"/>
        <SolidColorBrush x:Key="BrushScrollThumb" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushScrollTrack" Color="#050A14"/>
        <SolidColorBrush x:Key="BrushDisabledBg" Color="#1A2B44"/>
        <SolidColorBrush x:Key="BrushOk" Color="#34D399"/>
        <SolidColorBrush x:Key="BrushWarn" Color="#FBBF24"/>
        <SolidColorBrush x:Key="BrushErr" Color="#F87171"/>
        <SolidColorBrush x:Key="BrushGridHeaderBg" Color="#0B1528"/>
        <SolidColorBrush x:Key="BrushGridRowBg" Color="#091122"/>
        <SolidColorBrush x:Key="BrushGridAltRowBg" Color="#0C172B"/>
        <SolidColorBrush x:Key="BrushGridSelectedBg" Color="#1E3A5F"/>
        <SolidColorBrush x:Key="BrushGridLine" Color="#1A2B44"/>
        <SolidColorBrush x:Key="BrushDangerBg" Color="#26EF4444"/>
        <SolidColorBrush x:Key="BrushDangerBorder" Color="#66EF4444"/>
        <SolidColorBrush x:Key="BrushDangerFg" Color="#FCA5A5"/>
        <SolidColorBrush x:Key="BrushRowSmsOnlyBg" Color="#2BEF4444"/>
        <SolidColorBrush x:Key="BrushRowSmsOnlyFg" Color="#FCA5A5"/>
        <SolidColorBrush x:Key="BrushRowSmsDefBg" Color="#24FBBF24"/>
        <SolidColorBrush x:Key="BrushRowSmsDefFg" Color="#FDE68A"/>
        <SolidColorBrush x:Key="BrushRowSecureBg" Color="#1E34D399"/>
        <SolidColorBrush x:Key="BrushRowSecureFg" Color="#6EE7B7"/>
        <SolidColorBrush x:Key="BrushRowUnregBg" Color="#1A94A3B8"/>
        <SolidColorBrush x:Key="BrushRowUnregFg" Color="#CBD5E1"/>

        <!-- Primary CTA Button -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushDisabledBg}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushSecBtnPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="SelectionBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="Margin" Value="0,0,14,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Border x:Name="Box" Width="16" Height="16" CornerRadius="4"
                                    Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1"
                                    Margin="0,0,7,0" VerticalAlignment="Center">
                                <Path x:Name="CheckMark" Data="M 3,7.5 L 6.2,10.5 L 11.5,4"
                                      Stroke="{DynamicResource BrushAccentFg}" StrokeThickness="2" Visibility="Collapsed"
                                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"/>
                            </Border>
                            <ContentPresenter VerticalAlignment="Center"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box" Property="Background" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="{DynamicResource BrushScrollTrack}"/>
            <Setter Property="Width" Value="10"/>
            <Setter Property="MinWidth" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border Background="{DynamicResource BrushScrollThumb}" CornerRadius="4" Margin="2"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Margin="0,0,0,10">
                                <TabPanel IsItemsHost="True" Margin="16,0"/>
                            </Border>
                            <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="16,9"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource BrushAccentTint}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER BAR -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock Text="Entra ID MFA SMS Deprecation Checker" FontSize="19" FontWeight="Bold" Foreground="{DynamicResource BrushTextPrimary}"/>
                    <TextBlock Text="Audit user authentication methods, identify SMS-default &amp; SMS-only accounts, and export color-coded XLSX reports." FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}" Margin="0,2,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="BtnThemeToggle" Height="36" Margin="0,0,10,0" Padding="12,6" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <Border Background="{DynamicResource BrushSecBtnBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="11,6" Margin="0,0,12,0" VerticalAlignment="Center">
                        <TextBlock Name="txtConnStatus" Text="Not Connected" Foreground="{DynamicResource BrushErr}" FontWeight="SemiBold" FontSize="12" VerticalAlignment="Center"/>
                    </Border>
                    <Button Name="btnConnect" Content="Connect to Entra ID" Style="{StaticResource PrimaryButton}" Height="36"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- KPI STRIP -->
        <Grid Grid.Row="1" Margin="16,12,16,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="14,10" Margin="0,0,6,0">
                <StackPanel>
                    <TextBlock Text="Total Users Scanned" FontSize="11.5" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold"/>
                    <TextBlock Name="kpiTotal" Text="0" FontSize="22" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" Margin="0,3,0,0"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="1" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="14,10" Margin="3,0,3,0">
                <StackPanel>
                    <TextBlock Text="SMS ONLY (High Risk)" FontSize="11.5" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold"/>
                    <TextBlock Name="kpiSmsOnly" Text="0" FontSize="22" FontWeight="Bold" Foreground="{DynamicResource BrushErr}" FontFamily="Cascadia Mono, Consolas" Margin="0,3,0,0"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="2" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="14,10" Margin="3,0,3,0">
                <StackPanel>
                    <TextBlock Text="SMS Default (Medium Risk)" FontSize="11.5" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold"/>
                    <TextBlock Name="kpiSmsDefault" Text="0" FontSize="22" FontWeight="Bold" Foreground="{DynamicResource BrushWarn}" FontFamily="Cascadia Mono, Consolas" Margin="0,3,0,0"/>
                </StackPanel>
            </Border>

            <Border Grid.Column="3" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="14,10" Margin="6,0,0,0">
                <StackPanel>
                    <TextBlock Text="Secure MFA (Compliant)" FontSize="11.5" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold"/>
                    <TextBlock Name="kpiSecure" Text="0" FontSize="22" FontWeight="Bold" Foreground="{DynamicResource BrushOk}" FontFamily="Cascadia Mono, Consolas" Margin="0,3,0,0"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- HORIZONTAL SECTION TABS -->
        <TabControl Grid.Row="2" Margin="0,0,0,8">
            <TabItem Header="MFA Audit Results">
                <Grid Margin="16,0,16,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Controls Bar -->
                    <Border Grid.Row="0" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,0,0,10">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <CheckBox Name="chkActiveOnly" Content="Active Users Only" IsChecked="True" Grid.Column="0"/>
                            <CheckBox Name="chkExcludeGuests" Content="Exclude Guest Users" IsChecked="True" Grid.Column="1"/>

                            <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
                                <TextBlock Text="Filter:" VerticalAlignment="Center" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold" Margin="0,0,8,0"/>
                                <TextBox Name="txtSearch" Width="260" Height="32" VerticalContentAlignment="Center"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- DataGrid Preview -->
                    <DataGrid Name="dgAudit" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                              Background="{DynamicResource BrushGridRowBg}" Foreground="{DynamicResource BrushTextSecondary}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1"
                              RowBackground="{DynamicResource BrushGridRowBg}" AlternatingRowBackground="{DynamicResource BrushGridAltRowBg}"
                              GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{DynamicResource BrushGridLine}"
                              HeadersVisibility="Column" RowHeaderWidth="0" SelectionMode="Single" FontSize="12">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="{DynamicResource BrushGridHeaderBg}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
                                <Setter Property="Padding" Value="10,7"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="Background" Value="Transparent"/>
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Padding" Value="8,5"/>
                                <Setter Property="Foreground" Value="{Binding Foreground, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="{DynamicResource BrushGridSelectedBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.CellStyle>
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Setter Property="Background" Value="{DynamicResource BrushGridRowBg}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding RiskCategory}" Value="SMS ONLY">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowSmsOnlyBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowSmsOnlyFg}"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding RiskCategory}" Value="SMS DEFAULT">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowSmsDefBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowSmsDefFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding RiskCategory}" Value="SECURE MFA">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowSecureBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowSecureFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding RiskCategory}" Value="UNREGISTERED">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowUnregBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowUnregFg}"/>
                                    </DataTrigger>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="{DynamicResource BrushGridSelectedBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                    </Trigger>
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

                    <!-- Export Path Bar -->
                    <Border Grid.Row="2" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,10,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Export Path:" VerticalAlignment="Center" Foreground="{DynamicResource BrushTextMuted}" FontWeight="SemiBold" Margin="0,0,10,0"/>
                            <TextBox Name="txtExportPath" Grid.Column="1" Height="32" VerticalContentAlignment="Center" FontFamily="Cascadia Mono, Consolas" FontSize="12"/>
                            <Button Name="btnBrowse" Content="Browse..." Grid.Column="2" Height="32" Padding="14,4" Margin="10,0,0,0"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem Header="Activity Log">
                <Border Margin="16,0,16,0" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="10">
                    <TextBox Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                             Background="Transparent" BorderThickness="0" Padding="0"
                             Foreground="{DynamicResource BrushTextSecondary}" FontFamily="Cascadia Mono, Consolas" FontSize="12"/>
                </Border>
            </TabItem>
        </TabControl>

        <!-- SLIM PROGRESS BAR -->
        <ProgressBar Name="pbStatus" Grid.Row="3" Height="3" Minimum="0" Maximum="100" Value="0"
                     Background="{DynamicResource BrushGridHeaderBg}" Foreground="{DynamicResource BrushAccent}" BorderThickness="0"/>

        <!-- SINGLE-ROW FOOTER -->
        <Border Grid.Row="4" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Grid.Column="0" Text="Logs stream to Activity Log · Export creates a color-coded XLSX workbook (CSV fallback if ImportExcel is missing)"
                           Foreground="{DynamicResource BrushTextMuted}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="btnScan" Content="Scan MFA Methods" Height="36" MinWidth="150" Margin="0,0,10,0"/>
                    <Button Name="btnExport" Content="Export Color-Coded XLSX" Style="{StaticResource PrimaryButton}" Height="36" MinWidth="195"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Read XAML
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($inputXaml))
$window = [System.Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $window

# Map UI Controls
$btnThemeToggle = $window.FindName("BtnThemeToggle")
$txtThemeLabel  = $window.FindName("TxtThemeLabel")
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

# Theme State + Painted Win11 Chrome Hook
$global:IsDarkTheme = Get-SystemIsDarkTheme
$global:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $window -IsDark $global:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true

$window.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $window -IsDark $global:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $global:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $global:IsSystemThemeAuto) {
        $window.Dispatcher.Invoke([Action]{
            $global:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $window -IsDark $global:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)

if ($btnThemeToggle) {
    $btnThemeToggle.Add_Click({
        $global:IsSystemThemeAuto = $false
        $global:IsDarkTheme = -not $global:IsDarkTheme
        Set-StandardGuiTheme -Window $window -IsDark $global:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $false
    })
}

# Default Export Path
$defaultPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "MFA_SMS_Deprecation_Report_$((Get-Date).ToString('ddMMyy_HHmmss')).xlsx")
$txtExportPath.Text = $defaultPath

# Initial Log
Write-GuiLog "Entra ID MFA SMS Deprecation Checker initialized."

# Run module and dependency verification when the GUI window renders
$window.Add_Loaded({
    $btnConnect.IsEnabled = $false
    $btnScan.IsEnabled = $false
    $txtConnStatus.Text = "Checking Prerequisites..."
    $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushWarn')
    
    $modulesOk = Assert-RequiredModules
    
    $btnConnect.IsEnabled = $true
    $btnScan.IsEnabled = $true
    if ($modulesOk) {
        $txtConnStatus.Text = "Not Connected"
        $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushErr')
    } else {
        $txtConnStatus.Text = "Prerequisites Missing"
        $txtConnStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushErr')
    }
})

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

# Window Closing & Exit Cleanup Handler
$window.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
    try {
        Disconnect-MgGraph -ErrorAction Stop | Out-Null
    } catch {}
    $global:GraphConnected = $false
})

# Show GUI Window with Cleanup Finally Block
try {
    [void]$window.ShowDialog()
} finally {
    try {
        Disconnect-MgGraph -ErrorAction Stop | Out-Null
    } catch {}
    $global:GraphConnected = $false
}
