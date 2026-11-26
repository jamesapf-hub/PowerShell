#Requires -Version 7.2
<#
.SYNOPSIS
    M365 Admin Console - Native PowerShell 7 WPF Desktop Dashboard.
.DESCRIPTION
    Launches a modern, dark-themed (Kinetic Command) administration interface to audit
    licensed Entra ID users, evaluate last login timelines, resolve SKUs, and export results.
#>

# 1. In-Memory Guard for Bundled Package Scripts
if (-not $PSScriptRoot -or $PSScriptRoot -match "^iex" -or $MyInvocation.MyCommand.Path -like "*iex*") {
    Write-Host "[ERROR] This console is a packaged bundle requiring local relative files (LicensePrices.csv, LicencedUsersSigninDate.ps1)." -ForegroundColor Red
    Write-Host "[ERROR] In-memory execution via 'iex (irm ...)' is not supported." -ForegroundColor Red
    Write-Host "Please download and extract the package locally, then run LaunchConsole.bat or .\Console.ps1." -ForegroundColor Yellow
    return
}

# Force STA (Single Threaded Apartment) Mode for WPF
if ($Host.Runspace.ApartmentState -ne 'STA') {
    Write-Host "Apartment state is not STA. Restarting script in STA mode..." -ForegroundColor Yellow
    $Arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-ApartmentState", "STA", "-File", $MyInvocation.MyCommand.Path)
    Start-Process pwsh.exe -ArgumentList $Arguments -NoNewWindow -Wait
    return
}

# 2. Import Required Windows Presentation Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Xml, System.Drawing

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
            BrushChromeBar      = '#CC070D19'
            BrushHairline       = '#1CFFFFFF'
            BrushCardBg         = '#B30D192E'
            BrushCardBorder     = '#26354D'
            BrushInputBg        = '#C7070D1A'
            BrushLogBg          = '#E6050A14'
            BrushTextPrimary    = '#F8FAFC'
            BrushTextSecondary  = '#CBD5E1'
            BrushTextMuted      = '#94A3B8'
            BrushAccent         = '#2DD4BF'
            BrushAccentHover    = '#5EEAD4'
            BrushAccentPressed  = '#14B8A6'
            BrushAccentFg       = '#042F2E'
            BrushAccentTint     = '#242DD4BF'
            BrushAccentBorder   = '#6B2DD4BF'
            BrushSecBtnBg       = '#14FFFFFF'
            BrushSecBtnHover    = '#24FFFFFF'
            BrushSecBtnPressed  = '#0DFFFFFF'
            BrushScrollThumb    = '#26354D'
            BrushScrollTrack    = '#050A14'
            BrushDisabledBg     = '#1A2B44'
            BrushOk             = '#34D399'
            BrushOkTint         = '#1F34D399'
            BrushWarn           = '#FBBF24'
            BrushErr            = '#F87171'
            BrushInfo           = '#38BDF8'
            BrushGridHeaderBg   = '#0B1528'
            BrushGridRowBg      = '#091122'
            BrushGridAltRowBg   = '#0C172B'
            BrushGridSelectedBg = '#1E3A5F'
            BrushGridLine       = '#1A2B44'
            BrushDangerBg       = '#26EF4444'
            BrushDangerBorder   = '#66EF4444'
            BrushDangerFg       = '#FCA5A5'
            BrushRowActiveBg    = '#1A34D399'
            BrushRowActiveFg    = '#6EE7B7'
            BrushRowWarnBg      = '#1AFBBF24'
            BrushRowWarnFg      = '#FDE68A'
            BrushRowInactBg     = '#1A2DD4BF'
            BrushRowInactFg     = '#5EEAD4'
            BrushRowCritBg      = '#26EF4444'
            BrushRowCritFg      = '#FCA5A5'
            BrushRowNeverBg     = '#1494A3B8'
            BrushRowNeverFg     = '#CBD5E1'
        }
    } else {
        $canvasBrush = $conv.ConvertFromString('#F8F9FF')
        $map = @{
            BrushChromeBar      = '#F2FFFFFF'
            BrushHairline       = '#190F172A'
            BrushCardBg         = '#FFFFFFFF'
            BrushCardBorder     = '#CBD5E1'
            BrushInputBg        = '#F8FAFC'
            BrushLogBg          = '#F1F5F9'
            BrushTextPrimary    = '#0F172A'
            BrushTextSecondary  = '#334155'
            BrushTextMuted      = '#64748B'
            BrushAccent         = '#2563EB'
            BrushAccentHover    = '#1D4ED8'
            BrushAccentPressed  = '#1E40AF'
            BrushAccentFg       = '#FFFFFF'
            BrushAccentTint     = '#DBEAFE'
            BrushAccentBorder   = '#592563EB'
            BrushSecBtnBg       = '#FFFFFF'
            BrushSecBtnHover    = '#F1F5F9'
            BrushSecBtnPressed  = '#E2E8F0'
            BrushScrollThumb    = '#CBD5E1'
            BrushScrollTrack    = '#F1F5F9'
            BrushDisabledBg     = '#E2E8F0'
            BrushOk             = '#059669'
            BrushOkTint         = '#D1FAE5'
            BrushWarn           = '#B45309'
            BrushErr            = '#DC2626'
            BrushInfo           = '#0284C7'
            BrushGridHeaderBg   = '#F1F5F9'
            BrushGridRowBg      = '#FFFFFF'
            BrushGridAltRowBg   = '#F8FAFC'
            BrushGridSelectedBg = '#DBEAFE'
            BrushGridLine       = '#E2E8F0'
            BrushDangerBg       = '#FEE2E2'
            BrushDangerBorder   = '#FCA5A5'
            BrushDangerFg       = '#B91C1C'
            BrushRowActiveBg    = '#D1FAE5'
            BrushRowActiveFg    = '#065F46'
            BrushRowWarnBg      = '#FEF3C7'
            BrushRowWarnFg      = '#92400E'
            BrushRowInactBg     = '#DBEAFE'
            BrushRowInactFg     = '#1E40AF'
            BrushRowCritBg      = '#FEE2E2'
            BrushRowCritFg      = '#991B1B'
            BrushRowNeverBg     = '#F1F5F9'
            BrushRowNeverFg     = '#334155'
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

# Dependency Check & Auto-Installation
function Assert-ModuleInstalled {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        [Parameter(Mandatory=$true)]
        [string]$FriendlyName
    )
    Write-Host "Checking for dependency: ${FriendlyName} (${ModuleName})..." -ForegroundColor Cyan
    $Module = Get-Module -ListAvailable -Name $ModuleName
    if (-not $Module) {
        Write-Host "[WARNING] ${FriendlyName} is missing. Attempting automatic installation..." -ForegroundColor Yellow
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Write-Host "Installing ${ModuleName} from PowerShell Gallery for current user..." -ForegroundColor Gray
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "[SUCCESS] ${FriendlyName} successfully installed." -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Failed to install ${FriendlyName}: $_" -ForegroundColor Red
            [System.Windows.MessageBox]::Show(
                "Failed to install required dependency: ${FriendlyName} (${ModuleName}).`n`nError: $_`n`nPlease run PowerShell and install it manually: Install-Module $ModuleName -Scope CurrentUser",
                "Dependency Installation Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            exit 1
        }
    } else {
        Write-Host "[OK] $FriendlyName is available." -ForegroundColor Green
    }
}

Assert-ModuleInstalled -ModuleName "Microsoft.Graph" -FriendlyName "Microsoft Graph SDK"

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

# 3. GUI Layout Definition (XAML - Dynamic System Dark/Light Theme)
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft 365 Licence Console" Height="780" Width="1180" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual" Background="#091122" Foreground="{DynamicResource BrushTextPrimary}" FontSize="12.5">
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
        <SolidColorBrush x:Key="BrushOkTint" Color="#1F34D399"/>
        <SolidColorBrush x:Key="BrushWarn" Color="#FBBF24"/>
        <SolidColorBrush x:Key="BrushErr" Color="#F87171"/>
        <SolidColorBrush x:Key="BrushInfo" Color="#38BDF8"/>
        <SolidColorBrush x:Key="BrushGridHeaderBg" Color="#0B1528"/>
        <SolidColorBrush x:Key="BrushGridRowBg" Color="#091122"/>
        <SolidColorBrush x:Key="BrushGridAltRowBg" Color="#0C172B"/>
        <SolidColorBrush x:Key="BrushGridSelectedBg" Color="#1E3A5F"/>
        <SolidColorBrush x:Key="BrushGridLine" Color="#1A2B44"/>
        <SolidColorBrush x:Key="BrushDangerBg" Color="#26EF4444"/>
        <SolidColorBrush x:Key="BrushDangerBorder" Color="#66EF4444"/>
        <SolidColorBrush x:Key="BrushDangerFg" Color="#FCA5A5"/>
        <SolidColorBrush x:Key="BrushRowActiveBg" Color="#1A34D399"/>
        <SolidColorBrush x:Key="BrushRowActiveFg" Color="#6EE7B7"/>
        <SolidColorBrush x:Key="BrushRowWarnBg" Color="#1AFBBF24"/>
        <SolidColorBrush x:Key="BrushRowWarnFg" Color="#FDE68A"/>
        <SolidColorBrush x:Key="BrushRowInactBg" Color="#1A2DD4BF"/>
        <SolidColorBrush x:Key="BrushRowInactFg" Color="#5EEAD4"/>
        <SolidColorBrush x:Key="BrushRowCritBg" Color="#26EF4444"/>
        <SolidColorBrush x:Key="BrushRowCritFg" Color="#FCA5A5"/>
        <SolidColorBrush x:Key="BrushRowNeverBg" Color="#1494A3B8"/>
        <SolidColorBrush x:Key="BrushRowNeverFg" Color="#CBD5E1"/>

        <Style x:Key="BrandBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
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

        <Style x:Key="SecBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="12,6"/>
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
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="SelectionBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
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

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="PasswordBox">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
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
                            <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1">
                                <TabPanel IsItemsHost="True" Margin="18,0"/>
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
            <Setter Property="Padding" Value="16,10"/>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER BAR -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="20,13">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="Microsoft 365 Licence Console" Foreground="{DynamicResource BrushTextPrimary}" FontSize="19" FontWeight="Bold"/>
                    <TextBlock Text="Audit licensed Entra ID accounts, inactive sign-ins, unassigned SKU waste, and downgrade opportunities." Foreground="{DynamicResource BrushTextMuted}" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="BtnThemeToggle" Style="{StaticResource SecBtn}" Height="32" Margin="0,0,10,0" Padding="11,5" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <Border Name="bdgSessionStatus" Background="{DynamicResource BrushSecBtnBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,6" VerticalAlignment="Center">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Name="txtStatusLabel" Text="Offline Session" Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" FontWeight="SemiBold" FontFamily="Cascadia Mono, Consolas"/>
                            <Border Width="1" Height="12" Background="{DynamicResource BrushCardBorder}" Margin="10,0"/>
                            <TextBlock Text="TLS 1.2" Foreground="{DynamicResource BrushOk}" FontSize="11" FontFamily="Cascadia Mono, Consolas" FontWeight="Bold"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ARCHETYPE 2: HORIZONTAL SECTION TABS -->
        <TabControl Name="tabMain" Grid.Row="1">
            <!-- TAB 1: OVERVIEW -->
            <TabItem Name="tabOverview" Header="Overview">
                <Grid Name="pageDashboard" Margin="20,14,20,12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- KPIs Row -->
                    <Grid Grid.Row="0" Margin="0,0,0,14">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Name="colKpiTotal" Width="*"/>
                            <ColumnDefinition Name="colKpiActive" Width="*"/>
                            <ColumnDefinition Name="colKpiInactive90d" Width="*"/>
                            <ColumnDefinition Name="colKpiInactive1yr" Width="*"/>
                            <ColumnDefinition Name="colKpiSavings" Width="1.2*"/>
                            <ColumnDefinition Name="colKpiPoolWaste" Width="1.2*"/>
                        </Grid.ColumnDefinitions>

                        <Border Grid.Column="0" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,8,0">
                            <StackPanel>
                                <TextBlock Text="TOTAL AUDITED" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" FontWeight="Bold"/>
                                <TextBlock Name="kpiTotal" Text="0" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="22" FontWeight="Bold" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Grid.Column="1" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,8,0">
                            <StackPanel>
                                <TextBlock Text="ACTIVE (&lt;=30d)" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" FontWeight="Bold"/>
                                <TextBlock Name="kpiActive" Text="0" Foreground="{DynamicResource BrushOk}" FontFamily="Cascadia Mono, Consolas" FontSize="22" FontWeight="Bold" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Grid.Column="2" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,8,0">
                            <StackPanel>
                                <TextBlock Text="INACTIVE (&gt;90d)" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" FontWeight="Bold"/>
                                <TextBlock Name="kpiInactive90d" Text="0" Foreground="{DynamicResource BrushWarn}" FontFamily="Cascadia Mono, Consolas" FontSize="22" FontWeight="Bold" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Grid.Column="3" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,10" Margin="0,0,8,0">
                            <StackPanel>
                                <TextBlock Text="CRITICAL (&gt;1yr)" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" FontWeight="Bold"/>
                                <TextBlock Name="kpiInactive1yr" Text="0" Foreground="{DynamicResource BrushErr}" FontFamily="Cascadia Mono, Consolas" FontSize="22" FontWeight="Bold" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Name="borderKpiSavings" Grid.Column="4" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,8" Margin="0,0,8,0">
                            <StackPanel>
                                <TextBlock Text="WASTED MONEY (&gt;6m)" Foreground="{DynamicResource BrushTextMuted}" FontSize="9.5" FontWeight="Bold"/>
                                <TextBlock Name="kpiSavings" Text="£0.00" Foreground="{DynamicResource BrushErr}" FontFamily="Cascadia Mono, Consolas" FontSize="15" FontWeight="Bold" Margin="0,1,0,0"/>
                                <TextBlock Text="POTENTIAL MONTHLY SAVINGS" Foreground="{DynamicResource BrushTextMuted}" FontSize="9.5" FontWeight="Bold" Margin="0,3,0,0"/>
                                <TextBlock Name="kpiMonthlySavings" Text="£0.00/mo" Foreground="{DynamicResource BrushOk}" FontFamily="Cascadia Mono, Consolas" FontSize="15" FontWeight="Bold" Margin="0,1,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Name="borderKpiPoolWaste" Grid.Column="5" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12,8">
                            <StackPanel>
                                <TextBlock Text="POOL WASTE (UNASSIGNED)" Foreground="{DynamicResource BrushTextMuted}" FontSize="9.5" FontWeight="Bold"/>
                                <TextBlock Name="kpiPoolWaste" Text="£0.00" Foreground="{DynamicResource BrushErr}" FontFamily="Cascadia Mono, Consolas" FontSize="15" FontWeight="Bold" Margin="0,1,0,0"/>
                                <TextBlock Name="kpiPoolWasteCount" Text="0 unused licenses" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" Margin="0,3,0,0" FontWeight="SemiBold"/>
                            </StackPanel>
                        </Border>
                    </Grid>

                    <!-- Activity Distribution Card -->
                    <Border Grid.Row="1" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,12">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="User Activity Distribution" Foreground="{DynamicResource BrushTextPrimary}" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                <StackPanel Name="panelActivityChart"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>

                    <!-- Overview Export Actions -->
                    <Grid Grid.Row="2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Inspect full SKU pool waste on Licences or filter accounts on User Directory." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button Name="btnGenEmailDash" Style="{StaticResource SecBtn}" Content="Generate Email Draft" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportCSVDash" Style="{StaticResource SecBtn}" Content="Export CSV" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportExcelDash" Style="{StaticResource SecBtn}" Content="Export Excel (XLSX)" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportPDFDash" Style="{StaticResource SecBtn}" Content="Export PDF Report" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportPNGDash" Style="{StaticResource SecBtn}" Content="Snapshot PNG" Height="34"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: LICENCES -->
            <TabItem Name="tabLicences" Header="Licences">
                <Border Margin="20,14,20,12" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="14">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Tenant Subscriptions &amp; Unassigned Licence Pool" Foreground="{DynamicResource BrushTextPrimary}" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>
                        <DataGrid Grid.Row="1" Name="gridUnassigned" AutoGenerateColumns="False" IsReadOnly="True"
                                  Background="{DynamicResource BrushGridRowBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Foreground="{DynamicResource BrushTextSecondary}"
                                  RowBackground="{DynamicResource BrushGridRowBg}" AlternatingRowBackground="{DynamicResource BrushGridAltRowBg}"
                                  HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{DynamicResource BrushGridLine}"
                                  RowHeaderWidth="0" SelectionMode="Single" FontSize="12" RowHeight="30">
                            <DataGrid.ColumnHeaderStyle>
                                <Style TargetType="DataGridColumnHeader">
                                    <Setter Property="Background" Value="{DynamicResource BrushGridHeaderBg}"/>
                                    <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
                                    <Setter Property="Padding" Value="10,7"/>
                                    <Setter Property="FontWeight" Value="Bold"/>
                                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                    <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
                                </Style>
                            </DataGrid.ColumnHeaderStyle>
                            <DataGrid.CellStyle>
                                <Style TargetType="DataGridCell">
                                    <Setter Property="BorderThickness" Value="0"/>
                                    <Setter Property="Padding" Value="8,4"/>
                                    <Setter Property="VerticalAlignment" Value="Center"/>
                                    <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="{DynamicResource BrushGridSelectedBg}"/>
                                            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                        </Trigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.CellStyle>
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Product License" Binding="{Binding SkuPartName}" Width="2*"/>
                                <DataGridTextColumn Header="Assigned / Total" Binding="{Binding AssignedText}" Width="1.2*">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="HorizontalAlignment" Value="Center"/>
                                            <Setter Property="VerticalAlignment" Value="Center"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Unassigned" Binding="{Binding UnassignedUnits}" Width="1*">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="HorizontalAlignment" Value="Center"/>
                                            <Setter Property="VerticalAlignment" Value="Center"/>
                                            <Setter Property="FontWeight" Value="Bold"/>
                                            <Setter Property="Foreground" Value="{Binding UnassignedBrush}"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Monthly Cost" Binding="{Binding MonthlyPriceText}" Width="1.1*">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="HorizontalAlignment" Value="Right"/>
                                            <Setter Property="VerticalAlignment" Value="Center"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                                <DataGridTextColumn Header="Wasted/mo" Binding="{Binding WastedCostText}" Width="1.1*">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="HorizontalAlignment" Value="Right"/>
                                            <Setter Property="VerticalAlignment" Value="Center"/>
                                            <Setter Property="FontWeight" Value="SemiBold"/>
                                            <Setter Property="Foreground" Value="{Binding WastedBrush}"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                </Border>
            </TabItem>

            <!-- TAB 3: USER DIRECTORY -->
            <TabItem Name="tabDirectory" Header="User Directory">
                <Grid Name="pageDirectory" Margin="20,12,20,12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Filters & Search -->
                    <Grid Grid.Row="0" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="260"/>
                        </Grid.ColumnDefinitions>

                        <!-- Preset Filters -->
                        <WrapPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                            <Button Name="btnFilterAll" Content="All Users" Padding="10,5" Background="{DynamicResource BrushAccent}" Foreground="{DynamicResource BrushAccentFg}" FontWeight="Bold" BorderThickness="0" Margin="0,0,5,5" Tag="all"/>
                            <Button Name="btnFilterActive" Content="Active (&lt;=30d)" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="active"/>
                            <Button Name="btnFilterLicGuests" Content="Licensed + Guests" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="licensed_guests"/>
                            <Button Name="btnFilterLicensed" Content="Licensed Only" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="licensed"/>
                            <Button Name="btnFilterGuests" Content="Guests Only" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="guests"/>
                            <Button Name="btnFilterUnlicensed" Content="Unlicensed" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="unlicensed"/>
                            <Button Name="btnFilterDisabled" Content="Disabled" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="disabled"/>
                            <Button Name="btnFilterInactive90d" Content="Inactive (&gt;90d)" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="inactive90d"/>
                            <Button Name="btnFilterNever" Content="Never Logged In" Padding="10,5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,5,5" Tag="never"/>
                        </WrapPanel>

                        <!-- Search Box -->
                        <Grid Grid.Column="1" VerticalAlignment="Top">
                            <TextBox Name="txtSearch" Height="32" Padding="8,5" FontSize="12" VerticalContentAlignment="Center"/>
                            <TextBlock Text="Search Name, UPN, SKU..." Foreground="{DynamicResource BrushTextMuted}" IsHitTestVisible="False" VerticalAlignment="Center" Margin="10,0,0,0" FontSize="11.5">
                                <TextBlock.Style>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding ElementName=txtSearch, Path=Text}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>
                    </Grid>

                    <!-- Additive Filter Toolbar -->
                    <Border Grid.Row="1" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="8,5" Margin="0,0,0,8">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <WrapPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="REFINE:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" FontWeight="Bold" FontFamily="Cascadia Mono, Consolas" VerticalAlignment="Center" Margin="0,0,8,0"/>

                                <TextBlock Text="Lic:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10.5" VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <Button Name="btnSubLicAll" Content="All" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushAccentTint}" Foreground="{DynamicResource BrushTextPrimary}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" Margin="0,0,3,0" Tag="all"/>
                                <Button Name="btnSubLicYes" Content="Yes" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="licensed"/>
                                <Button Name="btnSubLicNo" Content="No" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,8,0" Tag="unlicensed"/>

                                <Border Width="1" Height="14" Background="{DynamicResource BrushCardBorder}" Margin="0,0,8,0" VerticalAlignment="Center"/>

                                <TextBlock Text="Type:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10.5" VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <Button Name="btnSubTypeAll" Content="All" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushAccentTint}" Foreground="{DynamicResource BrushTextPrimary}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" Margin="0,0,3,0" Tag="all"/>
                                <Button Name="btnSubTypeMembers" Content="Members" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="members"/>
                                <Button Name="btnSubTypeGuests" Content="Guests" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="guests"/>
                                <Button Name="btnSubTypeLicGuest" Content="Lic+Guests" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,8,0" Tag="licensed_guests"/>

                                <Border Width="1" Height="14" Background="{DynamicResource BrushCardBorder}" Margin="0,0,8,0" VerticalAlignment="Center"/>

                                <TextBlock Text="Status:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10.5" VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <Button Name="btnSubAccAll" Content="All" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushAccentTint}" Foreground="{DynamicResource BrushTextPrimary}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" Margin="0,0,3,0" Tag="all"/>
                                <Button Name="btnSubAccEnabled" Content="Enabled" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="enabled"/>
                                <Button Name="btnSubAccDisabled" Content="Disabled" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,8,0" Tag="disabled"/>

                                <Border Width="1" Height="14" Background="{DynamicResource BrushCardBorder}" Margin="0,0,8,0" VerticalAlignment="Center"/>

                                <TextBlock Text="Sign-In:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10.5" VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <Button Name="btnSubActAll" Content="All" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushAccentTint}" Foreground="{DynamicResource BrushTextPrimary}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" Margin="0,0,3,0" Tag="all"/>
                                <Button Name="btnSubActActive" Content="&lt;=30d" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="active"/>
                                <Button Name="btnSubActInactive" Content="&gt;90d" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,3,0" Tag="inactive90d"/>
                                <Button Name="btnSubActNever" Content="Never" Padding="6,3" FontSize="10.5" Background="{DynamicResource BrushSecBtnBg}" Foreground="{DynamicResource BrushTextMuted}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Margin="0,0,8,0" Tag="never"/>

                                <Button Name="btnResetFilters" Content="Reset All" Padding="8,3" FontSize="10.5" Background="{DynamicResource BrushDangerBg}" Foreground="{DynamicResource BrushDangerFg}" BorderBrush="{DynamicResource BrushDangerBorder}" BorderThickness="1" Margin="4,0,0,0"/>
                            </WrapPanel>

                            <TextBlock Name="txtFilterCount" Grid.Column="1" Text="Showing 0 of 0 users" Foreground="{DynamicResource BrushAccent}" FontSize="11.5" FontWeight="SemiBold" VerticalAlignment="Center" Margin="10,0,0,0"/>
                        </Grid>
                    </Border>

                    <!-- Data Grid -->
                    <DataGrid Name="gridUsers" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="True"
                              Background="{DynamicResource BrushGridRowBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Foreground="{DynamicResource BrushTextSecondary}"
                              RowBackground="{DynamicResource BrushGridRowBg}" AlternatingRowBackground="{DynamicResource BrushGridAltRowBg}"
                              HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{DynamicResource BrushGridLine}"
                              RowHeaderWidth="0" SelectionMode="Single" FontSize="11.5">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="{DynamicResource BrushGridHeaderBg}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
                                <Setter Property="Padding" Value="10,8"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="Background" Value="Transparent"/>
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Padding" Value="6,4"/>
                                <Setter Property="Foreground" Value="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
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
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                <Setter Property="BorderBrush" Value="{DynamicResource BrushGridLine}"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Active (&lt;=30d)">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowActiveBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowActiveFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (30-90d)">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowWarnBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowWarnFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (90-365d)">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowInactBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowInactFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (&gt;1yr)">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowCritBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowCritFg}"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Never Logged In">
                                        <Setter Property="Background" Value="{DynamicResource BrushRowNeverBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushRowNeverFg}"/>
                                    </DataTrigger>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="{DynamicResource BrushGridSelectedBg}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="1.2*"/>
                            <DataGridTextColumn Header="User Principal Name (UPN)" Binding="{Binding UserPrincipalName}" Width="1.4*"/>
                            <DataGridTextColumn Header="Type" Binding="{Binding UserType}" Width="75">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="HorizontalAlignment" Value="Center"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding UserType}" Value="Guest">
                                                <Setter Property="Foreground" Value="{DynamicResource BrushAccent}"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="Licenses" Binding="{Binding AssignedLicenses}" Width="1.8*"/>
                            <DataGridTextColumn Header="Last Sign-In" Binding="{Binding LastSignInDate}" Width="1.1*"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="120"/>
                            <DataGridTextColumn Header="Account Status" Binding="{Binding AccountStatusText}" Width="110">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="HorizontalAlignment" Value="Center"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="Wasted Cost" Binding="{Binding WastedCostText}" Width="95"/>
                            <DataGridTextColumn Header="Monthly Savings" Binding="{Binding MonthlySavingsText}" Width="95"/>
                            <DataGridTextColumn Header="Recommendation" Binding="{Binding Recommendation}" Width="1.4*"/>
                            <DataGridTextColumn Header="Verification" Binding="{Binding VerificationText}" Width="110"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Directory Export Actions -->
                    <Grid Grid.Row="3" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Results represent audited Entra ID accounts and subscriptions." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button Name="btnGenEmail" Style="{StaticResource SecBtn}" Content="Generate Email Draft" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportCSV" Style="{StaticResource SecBtn}" Content="Export CSV" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportExcel" Style="{StaticResource SecBtn}" Content="Export Excel (XLSX)" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportPDF" Style="{StaticResource SecBtn}" Content="Export PDF Report" Height="34" Margin="0,0,8,0"/>
                            <Button Name="btnExportPNG" Style="{StaticResource SecBtn}" Content="Snapshot PNG" Height="34"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 4: ACTIVITY LOG -->
            <TabItem Name="tabActivityLog" Header="Activity Log">
                <Border Margin="20,14,20,12" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="12">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="System Activity Feed" Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" FontWeight="SemiBold" Margin="0,0,0,8"/>
                        <TextBox Grid.Row="1" Name="txtLog" Background="Transparent" Foreground="{DynamicResource BrushTextSecondary}" BorderThickness="0" Padding="0"
                                 FontFamily="Cascadia Mono, Consolas" FontSize="11.5" IsReadOnly="True" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
                    </Grid>
                </Border>
            </TabItem>

            <!-- TAB 5: CONNECTION -->
            <TabItem Name="tabConnection" Header="Connection">
                <Grid Name="pageConnection" Margin="20,12,20,12">
                    <!-- Offline Welcome Screen -->
                    <ScrollViewer Name="panelOfflineWelcome" Visibility="Visible" VerticalScrollBarVisibility="Auto">
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Width="760" Margin="0,8,0,8">
                            <TextBlock Text="Microsoft 365 Licence Console" Foreground="{DynamicResource BrushTextPrimary}" FontSize="21" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,4"/>
                            <TextBlock Text="Select your operational role and authentication method to audit tenant users." Foreground="{DynamicResource BrushTextMuted}" FontSize="13" HorizontalAlignment="Center" Margin="0,0,0,18"/>

                            <!-- Step 1: Role Selection Cards -->
                            <TextBlock Text="STEP 1: SELECT YOUR OPERATING ROLE" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="11" FontWeight="Bold" Margin="0,0,0,8" HorizontalAlignment="Center"/>
                            <UniformGrid Columns="2" Margin="0,0,0,18" Width="620">
                                <Button Name="btnRoleSelectServiceDesk" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Padding="14" Margin="0,0,8,0" Cursor="Hand" Height="92">
                                    <Button.Style>
                                        <Style TargetType="Button">
                                            <Setter Property="Template">
                                                <Setter.Value>
                                                    <ControlTemplate TargetType="Button">
                                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                                            <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                                        </Border>
                                                    </ControlTemplate>
                                                </Setter.Value>
                                            </Setter>
                                        </Style>
                                    </Button.Style>
                                    <StackPanel>
                                        <TextBlock Text="[Desk] Service Desk" Foreground="{DynamicResource BrushTextPrimary}" FontSize="15.5" FontWeight="Bold" HorizontalAlignment="Center"/>
                                        <TextBlock Text="Standard view: licences, statuses &amp; accounts. Hides all financials." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" TextAlignment="Center" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </Button>

                                <Button Name="btnRoleSelectSales" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" Padding="14" Margin="8,0,0,0" Cursor="Hand" Height="92">
                                    <Button.Style>
                                        <Style TargetType="Button">
                                            <Setter Property="Template">
                                                <Setter.Value>
                                                    <ControlTemplate TargetType="Button">
                                                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                                            <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                                        </Border>
                                                    </ControlTemplate>
                                                </Setter.Value>
                                            </Setter>
                                        </Style>
                                    </Button.Style>
                                    <StackPanel>
                                        <TextBlock Text="[$] Sales &amp; Business" Foreground="{DynamicResource BrushTextPrimary}" FontSize="15.5" FontWeight="Bold" HorizontalAlignment="Center"/>
                                        <TextBlock Text="Financial view: savings, wasted costs, and downgrade recommendations." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" TextAlignment="Center" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </Button>
                            </UniformGrid>

                            <!-- Step 2: Connection Panel -->
                            <StackPanel Name="panelConnectionMethods" Visibility="Collapsed" HorizontalAlignment="Center" Width="760">
                                <TextBlock Text="STEP 2: CHOOSE CONNECTION METHOD" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="11" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>

                                <UniformGrid Columns="3" Margin="0,0,0,14">
                                    <!-- Option A: Interactive M365 Sign-In -->
                                    <Border Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,8,0">
                                        <Grid>
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="Interactive Sign-In" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13.5" FontWeight="Bold" Margin="0,0,0,6"/>
                                            <TextBlock Grid.Row="1" Text="Official Microsoft OAuth login prompt. Supports MFA and SSO." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" Margin="0,0,0,12"/>

                                            <StackPanel Grid.Row="2">
                                                <TextBlock Text="Tenant (Optional):" Foreground="{DynamicResource BrushTextMuted}" FontSize="10.5" Margin="0,0,0,3"/>
                                                <TextBox Name="txtUserTenant" Padding="6,5" Margin="0,0,0,8" FontSize="11.5"/>
                                                <Button Name="btnConnectM365" Style="{StaticResource BrandBtn}" Content="M365 Sign-In" Height="34" FontSize="12"/>
                                            </StackPanel>
                                        </Grid>
                                    </Border>

                                    <!-- Option B: App Client Secret -->
                                    <Border Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="4,0,4,0">
                                        <Grid>
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="App Client Secret" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13.5" FontWeight="Bold" Margin="0,0,0,6"/>
                                            <TextBlock Grid.Row="1" Text="Automated connection using Entra App Registration client credentials." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" Margin="0,0,0,10"/>

                                            <StackPanel Grid.Row="2">
                                                <TextBlock Text="Tenant ID/Domain:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" Margin="0,0,0,2"/>
                                                <TextBox Name="txtAppTenant" Padding="5,4" Margin="0,0,0,4" FontSize="11"/>
                                                <TextBlock Text="Client ID:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" Margin="0,0,0,2"/>
                                                <TextBox Name="txtAppClient" Padding="5,4" Margin="0,0,0,4" FontSize="11"/>
                                                <TextBlock Text="Client Secret:" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" Margin="0,0,0,2"/>
                                                <PasswordBox Name="txtAppSecret" Padding="5,4" Margin="0,0,0,8" FontSize="11"/>
                                                <Button Name="btnConnectApp" Style="{StaticResource BrandBtn}" Content="Connect App" Height="34" FontSize="12"/>
                                            </StackPanel>
                                        </Grid>
                                    </Border>

                                    <!-- Option C: Import CSV -->
                                    <Border Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="8,0,0,0">
                                        <Grid>
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <TextBlock Grid.Row="0" Text="Offline CSV Import" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13.5" FontWeight="Bold" Margin="0,0,0,6"/>
                                            <TextBlock Grid.Row="1" Text="Load previously exported user reports directly into the console." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" Margin="0,0,0,12"/>

                                            <Button Grid.Row="2" Name="btnBrowseCSV" Style="{StaticResource SecBtn}" Height="86">
                                                <StackPanel VerticalAlignment="Center">
                                                    <TextBlock Text="[CSV]" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="15" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,4"/>
                                                    <TextBlock Text="Browse CSV File" Foreground="{DynamicResource BrushTextPrimary}" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                                                    <TextBlock Text="Click to select file" Foreground="{DynamicResource BrushTextMuted}" FontSize="10" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                                                </StackPanel>
                                            </Button>
                                        </Grid>
                                    </Border>
                                </UniformGrid>

                                <!-- Demo Sandbox Option -->
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,4,0,0">
                                    <TextBlock Text="Need to test without credentials?" Foreground="{DynamicResource BrushTextMuted}" FontSize="12" VerticalAlignment="Center"/>
                                    <Button Name="btnDemoMode" Background="Transparent" BorderThickness="0" Foreground="{DynamicResource BrushAccent}" Content="Launch Demo Sandbox Mode" FontWeight="Bold" Cursor="Hand" Margin="8,0,0,0">
                                        <Button.Style>
                                            <Style TargetType="Button">
                                                <Setter Property="Template">
                                                    <Setter.Value>
                                                        <ControlTemplate TargetType="Button">
                                                            <TextBlock Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontWeight="{TemplateBinding FontWeight}" Cursor="{TemplateBinding Cursor}" TextDecorations="Underline"/>
                                                        </ControlTemplate>
                                                    </Setter.Value>
                                                </Setter>
                                            </Style>
                                        </Button.Style>
                                    </Button>
                                </StackPanel>
                            </StackPanel>
                        </StackPanel>
                    </ScrollViewer>

                    <!-- Connected Session Info Screen -->
                    <StackPanel Name="panelConnectedSession" Visibility="Collapsed" HorizontalAlignment="Center" VerticalAlignment="Center" Width="460">
                        <Border Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" CornerRadius="8" Padding="24">
                            <StackPanel>
                                <TextBlock Text="Active Session Established" Foreground="{DynamicResource BrushTextPrimary}" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,4"/>
                                <TextBlock Text="You are successfully authenticated and auditing the tenant." Foreground="{DynamicResource BrushTextMuted}" FontSize="12" HorizontalAlignment="Center" Margin="0,0,0,18"/>

                                <Grid Margin="0,6">
                                    <TextBlock Text="Tenant Domain:" Foreground="{DynamicResource BrushTextMuted}" FontSize="12"/>
                                    <TextBlock Name="txtSessionTenant" Text="N/A" Foreground="{DynamicResource BrushTextPrimary}" FontSize="12" FontWeight="Bold" HorizontalAlignment="Right"/>
                                </Grid>
                                <Grid Margin="0,6">
                                    <TextBlock Text="Connected As:" Foreground="{DynamicResource BrushTextMuted}" FontSize="12"/>
                                    <TextBlock Name="txtSessionUser" Text="N/A" Foreground="{DynamicResource BrushTextPrimary}" FontSize="12" HorizontalAlignment="Right" TextTrimming="CharacterEllipsis" MaxWidth="250"/>
                                </Grid>
                                <Grid Margin="0,6">
                                    <TextBlock Text="Uptime:" Foreground="{DynamicResource BrushTextMuted}" FontSize="12"/>
                                    <TextBlock Name="txtSessionUptime" Text="00:00:00" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="12" HorizontalAlignment="Right" FontWeight="Bold"/>
                                </Grid>
                                <Grid Margin="0,6,0,18">
                                    <TextBlock Text="Mode:" Foreground="{DynamicResource BrushTextMuted}" FontSize="12"/>
                                    <TextBlock Name="txtSessionMode" Text="N/A" Foreground="{DynamicResource BrushTextPrimary}" FontSize="12" HorizontalAlignment="Right"/>
                                </Grid>

                                <Button Name="btnDisconnect" Style="{StaticResource SecBtn}" Content="Disconnect Session" Height="38" FontSize="13"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- SINGLE-ROW FOOTER -->
        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Grid.Column="0" Text="Logs: C:\Logs\UserLicenceCheck · Use horizontal tabs to switch between Overview, Licences, Directory, Logs, and Connection"
                           Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="btnFooterActivity" Style="{StaticResource SecBtn}" Content="View Activity Log" Height="34" Margin="0,0,10,0"/>
                    <Button Name="btnRunScript" Style="{StaticResource BrandBtn}" Content="Refresh Report" Height="34" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

function Find-WpfChildByName {
    param(
        $Parent,
        [string]$Name
    )
    if ($null -eq $Parent) { return $null }
    if ($Parent.Name -eq $Name) { return $Parent }
    
    try {
        $Children = [System.Windows.LogicalTreeHelper]::GetChildren($Parent)
        foreach ($Child in $Children) {
            if ($Child -is [System.Windows.DependencyObject]) {
                $Found = Find-WpfChildByName -Parent $Child -Name $Name
                if ($null -ne $Found) { return $Found }
            }
        }
    } catch {}
    
    return $null
}

# 4. XML Parsing and WPF Window Binding
$Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($Xaml))
$Window = [System.Windows.Markup.XamlReader]::Load($Reader)
Initialize-StandardGuiWindow -Window $Window

# Dynamically link named XAML elements to local PowerShell variables prefixed with '$wpf_'
[xml]$Xml = $Xaml
$Xml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Val = $Window.FindName($_.Name)
    if ($null -eq $Val) {
        $Val = Find-WpfChildByName -Parent $Window -Name $_.Name
    }
    Set-Variable -Name "wpf_$($_.Name)" -Value $Val -Scope Script
}

function Update-ConsoleThemeElements {
    if ($Script:UnassignedLicenses -and $Window.Resources) {
        $okBrush   = $Window.Resources['BrushOk']
        $errBrush  = $Window.Resources['BrushErr']
        $muteBrush = $Window.Resources['BrushTextMuted']
        foreach ($item in $Script:UnassignedLicenses) {
            if ($item.PSObject.Properties['UnassignedTone']) {
                $item.UnassignedBrush = switch ($item.UnassignedTone) { 'Ok' { $okBrush } 'Err' { $errBrush } default { $muteBrush } }
            }
            if ($item.PSObject.Properties['WastedTone']) {
                $item.WastedBrush = switch ($item.WastedTone) { 'Err' { $errBrush } 'Ok' { $okBrush } default { $muteBrush } }
            }
        }
        if ($wpf_gridUnassigned) {
            try { $wpf_gridUnassigned.Items.Refresh() } catch {}
        }
    }
}

# Theme State + Painted Win11 Chrome Hook
$Script:IsDarkTheme = Get-SystemIsDarkTheme
$Script:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $wpf_TxtThemeLabel -IsSystemAuto $true

$Window.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $wpf_TxtThemeLabel -IsSystemAuto $Script:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $Script:IsSystemThemeAuto) {
        $Window.Dispatcher.Invoke([Action]{
            $Script:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $wpf_TxtThemeLabel -IsSystemAuto $true
            Update-ConsoleThemeElements
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)
$Window.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
})

if ($wpf_BtnThemeToggle) {
    $wpf_BtnThemeToggle.Add_Click({
        $Script:IsSystemThemeAuto = $false
        $Script:IsDarkTheme = -not $Script:IsDarkTheme
        Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $wpf_TxtThemeLabel -IsSystemAuto $false
        Update-ConsoleThemeElements
    })
}

# 5. Global State & Mocks
$Script:SessionMode = ""
$Script:CurrentData = $null
$Script:GridRows = $null
$Script:ActiveFilter = "all"
$Script:ActivePreset = "all"
$Script:FilterLicense = "all"
$Script:FilterAccount = "all"
$Script:FilterUserType = "all"
$Script:FilterActivity = "all"
$Script:SearchQuery = ""
$Script:SessionSeconds = 0
$Script:UserRole = ""
$Script:ExcludeAccountPatterns = if (-not [string]::IsNullOrWhiteSpace($env:M365_EXCLUDE_UPN_PATTERNS)) {
    @($env:M365_EXCLUDE_UPN_PATTERNS -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
} else {
    @("*msp-verify*", "*admin-verify*")
}

function Get-MockDateString($DaysAgo) {
    return (Get-Date).AddDays(-$DaysAgo).ToString("yyyy-MM-ddTHH:mm:ssZ")
}

$Script:MockData = @(
    [PSCustomObject]@{ DisplayName = "Alex Wilber"; UserPrincipalName = "AlexW@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium, Microsoft 365 Copilot"; LastSignInDate = (Get-MockDateString 2); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Adele Vance"; UserPrincipalName = "AdeleV@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 15); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Pradeep Gupta"; UserPrincipalName = "PradeepG@contoso.com"; AssignedLicenses = "Power BI Pro, Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 390); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Megan Bowen"; UserPrincipalName = "MeganB@contoso.com"; AssignedLicenses = "Microsoft 365 E3"; LastSignInDate = (Get-MockDateString 105); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Joni Sherman"; UserPrincipalName = "JoniS@contoso.com"; AssignedLicenses = "Microsoft 365 Business Basic"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $false; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Lynne Robbins"; UserPrincipalName = "LynneR@contoso.com"; AssignedLicenses = "Microsoft 365 F3, Exchange Online (Plan 2)"; LastSignInDate = (Get-MockDateString 410); AccountEnabled = $false; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Isaiah Langer"; UserPrincipalName = "IsaiahL@contoso.com"; AssignedLicenses = "Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 95); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Lidia Holloway"; UserPrincipalName = "LidiaH@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium, Power BI Pro"; LastSignInDate = (Get-MockDateString 0); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Grady Archie"; UserPrincipalName = "GradyA@contoso.com"; AssignedLicenses = "Microsoft 365 Business Standard"; LastSignInDate = (Get-MockDateString 195); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Patti Fernandez"; UserPrincipalName = "PattiF@contoso.com"; AssignedLicenses = "Microsoft Teams Exploratory"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Nestor Wilke"; UserPrincipalName = "NestorW@contoso.com"; AssignedLicenses = "Microsoft 365 Copilot, Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 35); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Diego Siciliani"; UserPrincipalName = "DiegoS@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 18); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Test Admin Account"; UserPrincipalName = "admin.test@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 5); AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Demo Admin User"; UserPrincipalName = "jp@contoso.com"; AssignedLicenses = "Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 12); AccountEnabled = $true; UserType = "Member" }
    # Unlicensed Members
    [PSCustomObject]@{ DisplayName = "Sales Shared Mailbox"; UserPrincipalName = "sales@contoso.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Conference Room Alpha"; UserPrincipalName = "room.alpha@contoso.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $true; UserType = "Member" }
    [PSCustomObject]@{ DisplayName = "Former Contractor (Disabled)"; UserPrincipalName = "contractor1@contoso.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = (Get-MockDateString 280); AccountEnabled = $false; UserType = "Member" }
    # Guest Accounts (Unlicensed & Licensed)
    [PSCustomObject]@{ DisplayName = "David Miller (Vendor Partner)"; UserPrincipalName = "david_miller_fabrikam.com#EXT#@contoso.onmicrosoft.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = (Get-MockDateString 8); AccountEnabled = $true; UserType = "Guest" }
    [PSCustomObject]@{ DisplayName = "Sarah Jenkins (External Auditor)"; UserPrincipalName = "sarah.j_auditco.com#EXT#@contoso.onmicrosoft.com"; AssignedLicenses = "Microsoft Fabric (Free)"; LastSignInDate = (Get-MockDateString 22); AccountEnabled = $true; UserType = "Guest" }
    [PSCustomObject]@{ DisplayName = "Elena Rostova (Guest Contractor)"; UserPrincipalName = "elena_freelance.org#EXT#@contoso.onmicrosoft.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = (Get-MockDateString 120); AccountEnabled = $true; UserType = "Guest" }
    [PSCustomObject]@{ DisplayName = "Inactive Guest Vendor"; UserPrincipalName = "oldguest_partner.com#EXT#@contoso.onmicrosoft.com"; AssignedLicenses = "None (Unlicensed)"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $false; UserType = "Guest" }
)

$Script:ActivityColors = @{
    "Active (<=30d)"          = "BrushOk"
    "Inactive (30-90d)"        = "BrushWarn"
    "Inactive (90-365d)"       = "BrushAccent"
    "Inactive (>1yr)"          = "BrushErr"
    "Never Logged In"          = "BrushTextMuted"
    "Requires Entra ID P1/P2" = "BrushInfo"
}

# Determine script directory
$Script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { try { Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction Stop } catch { $null } }
if (-not $Script:ScriptDir) { $Script:ScriptDir = (Get-Location).Path }

function Initialize-SkuPrices {
    $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
    
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
Microsoft Fabric (Free),0.00,a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
Power BI (free),0.00,a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
Power BI Pro,8.20,f8a1db68-be16-40ed-86d5-cb42ce701560
Power BI Pro,8.20,f8cdef31-a31e-4b4a-93e4-5f571e91255a
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
            Write-Host "Warning: Could not create default LicensePrices.csv: $_" -ForegroundColor Yellow
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
                # Fix a403ebcc incorrect mapping to Power BI Pro (Fabric Free is £0.00)
                if ($line -like "*a403ebcc-fae0-4ca2-8c8c-7a907fd6c235*") {
                    $CleanLines.Add("Microsoft Fabric (Free),0.00,a403ebcc-fae0-4ca2-8c8c-7a907fd6c235")
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

        try {
            $CurrentCsv = Get-Content -Path $PricesFile -Raw
            if ($CurrentCsv -notmatch "f8a1db68-be16-40ed-86d5-cb42ce701560") {
                Add-Content -Path $PricesFile -Value "`nPower BI Pro,8.20,f8a1db68-be16-40ed-86d5-cb42ce701560" -Encoding utf8
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
            Write-Host "Error loading LicensePrices.csv: $_" -ForegroundColor Red
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
            "Microsoft Fabric (Free)"            = 0.00
            "Power BI (free)"                    = 0.00
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

$Script:SkuPrices = Initialize-SkuPrices

# 6. Dispatcher Timer for Session Uptime
$Script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$Script:Timer.Interval = [TimeSpan]::FromSeconds(1)
$Script:Timer.Add_Tick({
    $Script:SessionSeconds++
    $hours = '{0:d2}' -f [int][Math]::Floor($Script:SessionSeconds / 3600)
    $minutes = '{0:d2}' -f [int][Math]::Floor(($Script:SessionSeconds % 3600) / 60)
    $seconds = '{0:d2}' -f [int]($Script:SessionSeconds % 60)
    $wpf_txtSessionUptime.Text = "${hours}:${minutes}:${seconds}"
})

# 7. Helper Functions
function Clear-MgGraphCache {
    try { Disconnect-MgGraph -ErrorAction Stop } catch {}
    try {
        Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force -ErrorAction Stop
    } catch {}
    try {
        Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction Stop
    } catch {}
    try {
        if (Test-Path "$env:LOCALAPPDATA\.IdentityService") {
            Get-ChildItem "$env:LOCALAPPDATA\.IdentityService" -Filter "mg*" -Recurse -ErrorAction Stop | ForEach-Object {
                try {
                    Remove-Item $_.FullName -Force -Recurse -ErrorAction Stop
                } catch {}
            }
        }
    } catch {}
}

function Start-SessionTimer {
    $Script:SessionSeconds = 0
    $wpf_txtSessionUptime.Text = "00:00:00"
    $Script:Timer.Start()
}

function Stop-SessionTimer {
    $Script:Timer.Stop()
}

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
            foreach ($key in $Script:SkuPrices.Keys) {
                # Check for exact display name, start/end matches, or exact SkuId GUID match
                if ($Name -eq $key -or $Name.StartsWith($key) -or $key.StartsWith($Name) -or ($Name -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$" -and $Name -eq $key.ToLower())) {
                    $Sum += $Script:SkuPrices[$key]
                    $Matched = $true
                    break
                }
            }
            if (-not $Matched) {
                if ($Name -match "Free" -or $Name -match "Exploratory" -or $Name -like "None*" -or $Name -match "Trial" -or $Name -match "Store for Business" -or ($Name -match "Basic" -and $Name -match "Teams")) {
                    $Sum += 0.00
                } elseif ($Name -match "E5") { $Sum += 48.10 }
                elseif ($Name -match "E3") { $Sum += 31.70 }
                elseif ($Name -match "Business Premium") { $Sum += 18.10 }
                elseif ($Name -match "Business Standard") { $Sum += 10.30 }
                elseif ($Name -match "Business Basic") { $Sum += 4.90 }
                elseif ($Name -match "F3") { $Sum += 6.60 }
                elseif ($Name -match "Copilot") { $Sum += 24.70 }
                else { $Sum += 0.00 }
            }
        }
    }
    return $Sum
}

function Resolve-SkuName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Guid,
        [Parameter(Mandatory=$false)]
        $AssignedPlans = $null
    )
    if ([string]::IsNullOrWhiteSpace($Guid)) { return "None (Unlicensed)" }
    $g = $Guid.ToString().ToLower().Trim()
    
    # 1. Fast memory cache lookup
    if ($Script:SkuMap -and $Script:SkuMap.ContainsKey($g)) {
        return $Script:SkuMap[$g]
    }
    
    # 2. Check tenant's Subscribed SKUs (using SkuPartNumber!)
    if ($Script:TenantSubSkus) {
        $subMatch = $Script:TenantSubSkus | Where-Object { $_.SkuId -and $_.SkuId.ToString().ToLower() -eq $g } | Select-Object -First 1
        if ($subMatch) {
            $pNum = if ($subMatch.SkuPartNumber) { $subMatch.SkuPartNumber.ToString().Trim() } elseif ($subMatch.SkuPartName) { $subMatch.SkuPartName.ToString().Trim() } else { "" }
            if ($pNum) {
                if ($Script:SkuMap -and $Script:SkuMap.ContainsKey($pNum.ToLower())) {
                    $found = $Script:SkuMap[$pNum.ToLower()]
                    if ($Script:SkuMap) { $Script:SkuMap[$g] = $found }
                    return $found
                }
                if ($Script:SkuCacheByStringId -and $Script:SkuCacheByStringId.ContainsKey($pNum.ToUpper())) {
                    $found = $Script:SkuCacheByStringId[$pNum.ToUpper()]
                    if ($Script:SkuMap) { $Script:SkuMap[$g] = $found }
                    return $found
                }
                $clean = ($pNum -replace '_', ' ').Trim()
                $clean = (Get-Culture).TextInfo.ToTitleCase($clean.ToLower())
                if ($Script:SkuMap) { $Script:SkuMap[$g] = $clean }
                return $clean
            }
        }
    }
    
    # 3. Check fast global SKU caches (GUID, Service_Plan_Id, String_Id)
    if ($Script:SkuCacheByGuid -and $Script:SkuCacheByGuid.ContainsKey($g)) {
        $found = $Script:SkuCacheByGuid[$g]
        if ($Script:SkuMap) { $Script:SkuMap[$g] = $found }
        return $found
    }
    if ($Script:SkuCacheByPlanId -and $Script:SkuCacheByPlanId.ContainsKey($g)) {
        $found = $Script:SkuCacheByPlanId[$g]
        if ($Script:SkuMap) { $Script:SkuMap[$g] = $found }
        return $found
    }
    
    # 4. Check on-demand in local M365_SKU_Cache.csv
    $CacheFile = Join-Path "$env:SystemDrive\Logs\UserLicenceCheck" "M365_SKU_Cache.csv"
    if (Test-Path $CacheFile) {
        try {
            $matchedRow = Import-Csv $CacheFile | Where-Object { $_.GUID -eq $g -or $_.Service_Plan_Id -eq $g } | Select-Object -First 1
            if ($matchedRow) {
                $name = if ($matchedRow.Product_Display_Name) { $matchedRow.Product_Display_Name } else { $matchedRow.Service_Plans_Included_Friendly_Names }
                if ($name) {
                    $name = $name -replace ';', ','
                    if ($Script:SkuMap) { $Script:SkuMap[$g] = $name }
                    return $name
                }
            }
        } catch {}
    }
    
    # 5. Live Online Dynamic Query: Query Merill's repo and Microsoft CSV on-the-fly
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $onlineUrl = "https://raw.githubusercontent.com/merill/license/main/license.csv"
        $onlineCsvText = Invoke-RestMethod -Uri $onlineUrl -TimeoutSec 4 -ErrorAction Stop
        if ($onlineCsvText) {
            $onlineList = ConvertFrom-Csv -InputObject $onlineCsvText
            $onlineMatch = $onlineList | Where-Object { $_.GUID -eq $g -or $_.Service_Plan_Id -eq $g } | Select-Object -First 1
            if ($onlineMatch) {
                $name = if ($onlineMatch.Product_Display_Name) { $onlineMatch.Product_Display_Name } else { $onlineMatch.Service_Plans_Included_Friendly_Names }
                if ($name) {
                    $name = $name -replace ';', ','
                    if ($Script:SkuMap) { $Script:SkuMap[$g] = $name }
                    try {
                        Add-Content -Path $CacheFile -Value "`n$($onlineMatch.Product_Display_Name),$($onlineMatch.String_Id),$($onlineMatch.GUID),$($onlineMatch.Service_Plan_Name),$($onlineMatch.Service_Plan_Id),$($onlineMatch.Service_Plans_Included_Friendly_Names)" -Encoding utf8
                    } catch {}
                    return $name
                }
            }
        }
    } catch {}
    
    # 6. Fallback: User AssignedPlans service analysis
    if ($AssignedPlans) {
        $activePlans = $AssignedPlans | Where-Object { $_.CapabilityStatus -eq "Enabled" } | Select-Object -ExpandProperty Service -Unique
        if ($activePlans -and $activePlans.Count -gt 0) {
            $summary = ($activePlans | Select-Object -First 2) -join " / "
            $inferred = "Custom Plan ($summary)"
            if ($Script:SkuMap) { $Script:SkuMap[$g] = $inferred }
            return $inferred
        }
    }
    
    return "Unknown SKU ($g)"
}

function Update-UI {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action] {}
    )
}

function DoEvents() {
    Update-UI
}

function Get-TenantName {
    try {
        if ($null -ne $Script:wpf_txtSessionTenant) {
            $val = $Script:wpf_txtSessionTenant.Text
            if (-not [string]::IsNullOrEmpty($val)) { return $val.Trim() }
        }
        if ($null -ne $wpf_txtSessionTenant) {
            $val = $wpf_txtSessionTenant.Text
            if (-not [string]::IsNullOrEmpty($val)) { return $val.Trim() }
        }
    } catch {}
    return "Tenant"
}

function Log-ToTerminal($Message, $Type = "Info") {
    $Timestamp = (Get-Date).ToString("HH:mm:ss")
    $Prefix = "pwsh> "
    if ($Type -eq "Success") { $Prefix = "[SUCCESS] " }
    elseif ($Type -eq "Warning") { $Prefix = "[WARNING] " }
    elseif ($Type -eq "Error") { $Prefix = "[ERROR] " }
    elseif ($Type -eq "Command") { $Prefix = "pwsh> " }
    
    $Line = "[$Timestamp] $Prefix$Message`r`n"
    $wpf_txtLog.AppendText($Line)
    $wpf_txtLog.ScrollToEnd()
    DoEvents
}

# Show-RoleSelectorDialog is deprecated. Role selection is performed directly on the welcome panel.

function Set-DataGridColumnVisibility {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HeaderName,
        [Parameter(Mandatory=$true)]
        [bool]$IsVisible
    )
    if ($null -eq $wpf_gridUsers) { return }
    $Col = $wpf_gridUsers.Columns | Where-Object { $_.Header -eq $HeaderName }
    if ($Col) {
        $Col.Visibility = if ($IsVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
}

function Set-UnassignedGridColumnVisibility {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HeaderName,
        [Parameter(Mandatory=$true)]
        [bool]$IsVisible
    )
    if ($null -eq $wpf_gridUnassigned) { return }
    $Col = $wpf_gridUnassigned.Columns | Where-Object { $_.Header -eq $HeaderName }
    if ($Col) {
        $Col.Visibility = if ($IsVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
}

function Configure-RoleUI {

    if ($Script:UserRole -eq "ServiceDesk") {
        # Hide monetary KPIs
        if ($wpf_borderKpiSavings) { $wpf_borderKpiSavings.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_borderKpiPoolWaste) { $wpf_borderKpiPoolWaste.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_colKpiSavings) { $wpf_colKpiSavings.Width = New-Object System.Windows.GridLength(0) }
        if ($wpf_colKpiPoolWaste) { $wpf_colKpiPoolWaste.Width = New-Object System.Windows.GridLength(0) }
        
        # Hide DataGrid columns
        Set-DataGridColumnVisibility -HeaderName "Wasted Cost" -IsVisible $false
        Set-DataGridColumnVisibility -HeaderName "Monthly Savings" -IsVisible $false
        Set-DataGridColumnVisibility -HeaderName "Recommendation" -IsVisible $false
        
        # Hide PDF exporter
        if ($wpf_btnExportPDF) { $wpf_btnExportPDF.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_btnExportPDFDash) { $wpf_btnExportPDFDash.Visibility = [System.Windows.Visibility]::Collapsed }
        
        # Hide monetary columns in Unassigned Licenses Breakdown Grid
        Set-UnassignedGridColumnVisibility -HeaderName "Monthly Cost" -IsVisible $false
        Set-UnassignedGridColumnVisibility -HeaderName "Wasted/mo" -IsVisible $false
    } else {
        # Show monetary KPIs
        if ($wpf_borderKpiSavings) { $wpf_borderKpiSavings.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_borderKpiPoolWaste) { $wpf_borderKpiPoolWaste.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_colKpiSavings) { $wpf_colKpiSavings.Width = New-Object System.Windows.GridLength(1.2, [System.Windows.GridUnitType]::Star) }
        if ($wpf_colKpiPoolWaste) { $wpf_colKpiPoolWaste.Width = New-Object System.Windows.GridLength(1.2, [System.Windows.GridUnitType]::Star) }
        
        # Show DataGrid columns
        Set-DataGridColumnVisibility -HeaderName "Wasted Cost" -IsVisible $true
        Set-DataGridColumnVisibility -HeaderName "Monthly Savings" -IsVisible $true
        Set-DataGridColumnVisibility -HeaderName "Recommendation" -IsVisible $true
        
        # Show PDF exporter
        if ($wpf_btnExportPDF) { $wpf_btnExportPDF.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_btnExportPDFDash) { $wpf_btnExportPDFDash.Visibility = [System.Windows.Visibility]::Visible }
        
        # Show monetary columns in Unassigned Licenses Breakdown Grid
        Set-UnassignedGridColumnVisibility -HeaderName "Monthly Cost" -IsVisible $true
        Set-UnassignedGridColumnVisibility -HeaderName "Wasted/mo" -IsVisible $true
    }
}

function Show-VerificationSelectorDialog {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    
    $DialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exclude Verification Accounts" Height="360" Width="460"
        WindowStartupLocation="CenterOwner" Background="#091122" Foreground="{DynamicResource BrushTextPrimary}"
        ResizeMode="NoResize" ShowInTaskbar="False">
    <Window.Resources>
        <SolidColorBrush x:Key="BrushCanvas" Color="#091122"/>
        <SolidColorBrush x:Key="BrushLogBg" Color="#E6050A14"/>
        <SolidColorBrush x:Key="BrushCardBorder" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushTextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BrushTextSecondary" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="BrushTextMuted" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BrushAccent" Color="#2DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentHover" Color="#5EEAD4"/>
        <SolidColorBrush x:Key="BrushAccentFg" Color="#042F2E"/>
        <SolidColorBrush x:Key="BrushSecBtnBg" Color="#14FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnHover" Color="#24FFFFFF"/>
        <Style x:Key="DialogBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
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
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DialogSecBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Verification Accounts Detected" Foreground="{DynamicResource BrushAccent}" FontSize="15" FontWeight="Bold"/>
            <TextBlock Text="The following Test/Admin accounts were found in the list. Uncheck any accounts you want to remove from the exported file." 
                       Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" TextWrapping="Wrap" Margin="0,4,0,0"/>
        </StackPanel>
        
        <Border Grid.Row="1" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="panelCheckBoxes"/>
            </ScrollViewer>
        </Border>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnProceed" Style="{StaticResource DialogBtn}" Content="Proceed Export" Width="125" Height="32" IsDefault="True" Margin="0,0,10,0"/>
            <Button Name="btnCancel" Style="{StaticResource DialogSecBtn}" Content="Cancel" Width="84" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($DialogXaml))
    $DlgWindow = [System.Windows.Markup.XamlReader]::Load($Reader)
    $DlgWindow.Owner = $Window
    if ($Window -and $Window.FontFamily) { $DlgWindow.FontFamily = $Window.FontFamily }
    $dlgIsDark = if ($null -ne $Script:IsDarkTheme) { [bool]$Script:IsDarkTheme } else { Get-SystemIsDarkTheme }
    Set-StandardGuiTheme -Window $DlgWindow -IsDark $dlgIsDark
    $DlgWindow.Add_SourceInitialized({
        $curDark = if ($null -ne $Script:IsDarkTheme) { [bool]$Script:IsDarkTheme } else { Get-SystemIsDarkTheme }
        Set-StandardGuiTheme -Window $DlgWindow -IsDark $curDark
    })
    
    $Panel = $DlgWindow.FindName("panelCheckBoxes")
    $BtnProceed = $DlgWindow.FindName("btnProceed")
    $BtnCancel = $DlgWindow.FindName("btnCancel")
    
    $CheckBoxes = @()
    foreach ($user in $Users) {
        $CB = [System.Windows.Controls.CheckBox]::new()
        $CB.Content = "$($user.DisplayName) ($($user.UserPrincipalName))"
        $CB.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushTextPrimary')
        $CB.IsChecked = $true
        $CB.Margin = "0,4,0,4"
        $CB.Tag = $user.UserPrincipalName
        $Panel.Children.Add($CB)
        $CheckBoxes += $CB
    }
    
    $Script:DlgResult = $null
    $Script:ExcludedUPNs = @()
    
    $BtnProceed.Add_Click({
        $Script:DlgResult = "Proceed"
        $Script:ExcludedUPNs = foreach ($cb in $CheckBoxes) {
            if ($cb.IsChecked -eq $false) {
                $cb.Tag
            }
        }
        $DlgWindow.Close()
    })
    
    $BtnCancel.Add_Click({
        $Script:DlgResult = "Cancel"
        $DlgWindow.Close()
    })
    
    $DlgWindow.ShowDialog() | Out-Null
    
    return [PSCustomObject]@{
        Result       = $Script:DlgResult
        ExcludedUPNs = $Script:ExcludedUPNs
    }
}

function Show-EmailTemplateDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$EmailText
    )
    
    $DialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Generated Email Draft" Height="500" Width="650"
        WindowStartupLocation="CenterOwner" Background="#091122" Foreground="{DynamicResource BrushTextPrimary}"
        ResizeMode="NoResize" ShowInTaskbar="False">
    <Window.Resources>
        <SolidColorBrush x:Key="BrushCanvas" Color="#091122"/>
        <SolidColorBrush x:Key="BrushLogBg" Color="#E6050A14"/>
        <SolidColorBrush x:Key="BrushCardBorder" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushTextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BrushTextSecondary" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="BrushTextMuted" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BrushAccent" Color="#2DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentHover" Color="#5EEAD4"/>
        <SolidColorBrush x:Key="BrushAccentFg" Color="#042F2E"/>
        <SolidColorBrush x:Key="BrushSecBtnBg" Color="#14FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnHover" Color="#24FFFFFF"/>
        <SolidColorBrush x:Key="BrushOk" Color="#34D399"/>
        <Style x:Key="DialogBtn" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="15,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="5,0"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BrandDialogBtn" TargetType="Button" BasedOn="{StaticResource DialogBtn}">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="Bold"/>
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
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <StackPanel Margin="0,0,0,15">
            <TextBlock Text="GENERATED EMAIL DRAFT" Foreground="{DynamicResource BrushAccent}" FontSize="14" FontWeight="Bold"/>
            <TextBlock Text="Review the generated draft below. You can copy it directly to your clipboard." Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" Margin="0,3,0,0"/>
        </StackPanel>
        
        <!-- Scrollable TextBox -->
        <Border Grid.Row="1" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="6">
            <TextBox Name="txtEmailBody" Background="Transparent" Foreground="{DynamicResource BrushTextPrimary}" BorderThickness="0"
                     TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     FontSize="12" Padding="8"/>
        </Border>
        
        <!-- Actions -->
        <Grid Grid.Row="2" Margin="0,15,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <TextBlock Name="txtStatusMessage" Text="" Foreground="{DynamicResource BrushOk}" FontSize="12" VerticalAlignment="Center" FontWeight="SemiBold"/>
            
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnCopy" Style="{StaticResource BrandDialogBtn}" Content="Copy to Clipboard"/>
                <Button Name="btnClose" Style="{StaticResource DialogBtn}" Content="Close"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

    # Apply blur to parent window
    $WpfWindow = $Window
    $Blur = New-Object System.Windows.Media.Effects.BlurEffect
    $Blur.Radius = 15
    $WpfWindow.Effect = $Blur
    
    # Load dialog
    $Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($DialogXaml))
    $DlgWindow = [System.Windows.Markup.XamlReader]::Load($Reader)
    $DlgWindow.Owner = $WpfWindow
    if ($WpfWindow -and $WpfWindow.FontFamily) { $DlgWindow.FontFamily = $WpfWindow.FontFamily }
    $dlgIsDark = if ($null -ne $Script:IsDarkTheme) { [bool]$Script:IsDarkTheme } else { Get-SystemIsDarkTheme }
    Set-StandardGuiTheme -Window $DlgWindow -IsDark $dlgIsDark
    $DlgWindow.Add_SourceInitialized({
        $curDark = if ($null -ne $Script:IsDarkTheme) { [bool]$Script:IsDarkTheme } else { Get-SystemIsDarkTheme }
        Set-StandardGuiTheme -Window $DlgWindow -IsDark $curDark
    })
    
    # Bind dialog controls
    $TxtEmailBody = $DlgWindow.FindName("txtEmailBody")
    $BtnCopy = $DlgWindow.FindName("btnCopy")
    $BtnClose = $DlgWindow.FindName("btnClose")
    $TxtStatusMessage = $DlgWindow.FindName("txtStatusMessage")
    
    if ($null -ne $TxtEmailBody) {
        $TxtEmailBody.Text = $EmailText
    }
    
    # Click Events
    if ($null -ne $BtnCopy) {
        $BtnCopy.Add_Click({
            if ($null -ne $TxtEmailBody) {
                [System.Windows.Clipboard]::SetText($TxtEmailBody.Text)
            }
            if ($null -ne $TxtStatusMessage) {
                $TxtStatusMessage.Text = "Copied to clipboard!"
            }
        })
    }
    
    if ($null -ne $BtnClose) {
        $BtnClose.Add_Click({
            $DlgWindow.Close()
        })
    }
    
    $DlgWindow.Add_Closed({
        $WpfWindow.Effect = $null
    })
    
    $DlgWindow.ShowDialog() | Out-Null
    
    # Remove blur
    $WpfWindow.Effect = $null
}

function Invoke-AuditReport {
    $wpf_btnRunScript.IsEnabled = $false
    $wpf_btnRunScript.Content = "Executing audit..."
    Update-UI
    
    Log-ToTerminal "pwsh> .\\LicencedUsersSigninDate.ps1" "Command"
    Update-UI
    
    Start-Sleep -Milliseconds 400
    Log-ToTerminal "Checking Microsoft Graph connection status..." "Info"
    Update-UI
    
    Start-Sleep -Milliseconds 400
    Log-ToTerminal "Connection verified. Fetching tenant organization metadata..." "Info"
    Update-UI
    
    if ($Script:SessionMode -eq "Demo") {
        Start-Sleep -Milliseconds 600
        Log-ToTerminal "Querying Entra licensed users (Simulated)..." "Info"
        Update-UI
        Start-Sleep -Milliseconds 400
        Log-ToTerminal "Resolving license SKU mappings..." "Info"
        Update-UI
        Start-Sleep -Milliseconds 400
        
        $GreenBrush = $Window.Resources['BrushOk']
        $RedBrush = $Window.Resources['BrushErr']
        $GrayBrush = $Window.Resources['BrushTextMuted']
        
        $Script:TotalPoolWaste = 186.70
        $Script:TotalUnassignedCount = 7
        $Script:UnassignedLicenses = @(
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 E5"; SkuId = "078d10ee-6995-4851-8043-334f610f49b3"; ActiveUnits = 10; ConsumedUnits = 8; UnassignedUnits = 2; UnassignedTone = "Ok"; WastedTone = "Err"; UnassignedBrush = $GreenBrush; WastedBrush = $RedBrush; MonthlyPrice = 48.10; MonthlyPriceText = "£48.10"; WastedCost = 96.20; WastedCostText = "£96.20"; AssignedText = "8 of 10" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Premium"; SkuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"; ActiveUnits = 25; ConsumedUnits = 20; UnassignedUnits = 5; UnassignedTone = "Ok"; WastedTone = "Err"; UnassignedBrush = $GreenBrush; WastedBrush = $RedBrush; MonthlyPrice = 18.10; MonthlyPriceText = "£18.10"; WastedCost = 90.50; WastedCostText = "£90.50"; AssignedText = "20 of 25" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Copilot"; SkuId = "639dec6b-bb19-468b-871c-c5c441c4b0cb"; ActiveUnits = 5; ConsumedUnits = 5; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 24.70; MonthlyPriceText = "£24.70"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "5 of 5" }
            [PSCustomObject]@{ SkuPartName = "Power BI Pro"; SkuId = "f8a1db68-be16-40ed-86d5-cb42ce701560"; ActiveUnits = 10; ConsumedUnits = 10; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 8.20; MonthlyPriceText = "£8.20"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "10 of 10" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Standard"; SkuId = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"; ActiveUnits = 15; ConsumedUnits = 15; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 10.30; MonthlyPriceText = "£10.30"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "15 of 15" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Basic"; SkuId = "bd251394-b1ed-487b-a1aa-ee198c62c938"; ActiveUnits = 30; ConsumedUnits = 30; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 4.90; MonthlyPriceText = "£4.90"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "30 of 30" }
            [PSCustomObject]@{ SkuPartName = "Office 365 E3"; SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"; ActiveUnits = 12; ConsumedUnits = 12; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 22.00; MonthlyPriceText = "£22.00"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "12 of 12" }
            [PSCustomObject]@{ SkuPartName = "Microsoft Teams Exploratory"; SkuId = "710779e8-3d4a-4c88-adb9-386c958d1fdf"; ActiveUnits = 50; ConsumedUnits = 50; UnassignedUnits = 0; UnassignedTone = "Muted"; WastedTone = "Muted"; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 0.00; MonthlyPriceText = "£0.00"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "50 of 50" }
        )
        $Users = $Script:MockData
        Process-UserData $Users
        Log-ToTerminal "Report generation complete. Successfully processed $($Users.Count) sandbox users." "Success"
    } else {
        # REAL AD HOC GRAPH RETRIEVAL
        try {
            Log-ToTerminal "Loading license SkuMap configurations..." "Info"
            Update-UI
            
            # Sku fallback dictionary
            $LocalSkuDict = @{
                "ab5128ae-2475-4d95-8c73-33f07d701bfc" = "Microsoft 365 Copilot"
                "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Microsoft Fabric (Free)"
                "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
                "f8cdef31-a31e-4b4a-93e4-5f571e91255a" = "Power BI Pro"
                "639dec6b-bb19-468b-871c-c5c441c4b0cb" = "Microsoft 365 Copilot"
                "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
                "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
                "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
                "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
                "61346032-1554-4736-b876-7d28d697f394" = "Microsoft 365 F3"
                "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87" = "Microsoft 365 F3"
                "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
                "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
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
                "POWER_BI_STANDARD"        = "Microsoft Fabric (Free)"
                "POWER_BI_PRO"             = "Power BI Pro"
            }
            
            $Script:SkuMap = @{}
            foreach ($key in $LocalSkuDict.Keys) {
                $Script:SkuMap[$key.ToLower()] = $LocalSkuDict[$key]
            }
            
            # Pre-populate friendly names from LicensePrices.csv
            $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
            if (Test-Path $PricesFile) {
                try {
                    $PricesCsv = Import-Csv $PricesFile
                    foreach ($row in $PricesCsv) {
                        if ($row.SkuId -and $row.Product) {
                            $Script:SkuMap[$row.SkuId.ToString().ToLower()] = $row.Product
                        }
                    }
                } catch {}
            }

            # Pre-load Global SKU Database FIRST so tenant SKUs and users resolve against 2,900+ official products
            $LogDir = "$env:SystemDrive\Logs\UserLicenceCheck"
            $CacheFile = Join-Path $LogDir "M365_SKU_Cache.csv"
            $CacheLoaded = $false
            $Script:SkuCacheByGuid = @{}
            $Script:SkuCacheByPlanId = @{}
            $Script:SkuCacheByStringId = @{}

            if (Test-Path $CacheFile) {
                $LastMod = (Get-Item $CacheFile).LastWriteTime
                if ($LastMod -gt (Get-Date).AddDays(-7)) {
                    Log-ToTerminal "Loading global SKU catalog from local cache..." "Info"
                    Update-UI
                    try {
                        Import-Csv $CacheFile | ForEach-Object {
                            if ($_.GUID -and $_.Product_Display_Name) {
                                $g = $_.GUID.ToString().ToLower().Trim()
                                $pName = $_.Product_Display_Name -replace ';', ','
                                $Script:SkuCacheByGuid[$g] = $pName
                                if (-not $Script:SkuMap.ContainsKey($g)) { $Script:SkuMap[$g] = $pName }
                            }
                            if ($_.Service_Plan_Id) {
                                $spId = $_.Service_Plan_Id.ToString().ToLower().Trim()
                                $spName = if ($_.Service_Plans_Included_Friendly_Names) { $_.Service_Plans_Included_Friendly_Names } else { $_.Product_Display_Name }
                                $Script:SkuCacheByPlanId[$spId] = $spName -replace ';', ','
                            }
                            if ($_.String_Id -and $_.Product_Display_Name) {
                                $sId = $_.String_Id.ToString().ToUpper().Trim()
                                $Script:SkuCacheByStringId[$sId] = $_.Product_Display_Name -replace ';', ','
                            }
                        }
                        $CacheLoaded = $true
                    } catch {}
                }
            }

            if (-not $CacheLoaded) {
                Log-ToTerminal "Fetching latest global SKU mapping database online..." "Info"
                Update-UI
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    $CsvText = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/merill/license/main/license.csv" -TimeoutSec 10
                    $OnlineSkus = ConvertFrom-Csv -InputObject $CsvText
                    foreach ($Sku in $OnlineSkus) {
                        if ($Sku.GUID -and $Sku.Product_Display_Name) {
                            $g = $Sku.GUID.ToString().ToLower().Trim()
                            $pName = $Sku.Product_Display_Name -replace ';', ','
                            $Script:SkuCacheByGuid[$g] = $pName
                            if (-not $Script:SkuMap.ContainsKey($g)) { $Script:SkuMap[$g] = $pName }
                        }
                        if ($Sku.Service_Plan_Id) {
                            $spId = $Sku.Service_Plan_Id.ToString().ToLower().Trim()
                            $spName = if ($Sku.Service_Plans_Included_Friendly_Names) { $Sku.Service_Plans_Included_Friendly_Names } else { $Sku.Product_Display_Name }
                            $Script:SkuCacheByPlanId[$spId] = $spName -replace ';', ','
                        }
                        if ($Sku.String_Id -and $Sku.Product_Display_Name) {
                            $sId = $Sku.String_Id.ToString().ToUpper().Trim()
                            $Script:SkuCacheByStringId[$sId] = $Sku.Product_Display_Name -replace ';', ','
                        }
                    }
                    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
                    $CsvText | Out-File -FilePath $CacheFile -Force -Encoding utf8
                    Log-ToTerminal "Successfully updated global SKU cache database." "Success"
                    Update-UI
                } catch {
                    Log-ToTerminal "Could not reach online SKU repository. Using local default fallbacks." "Warning"
                    Update-UI
                }
            }
            
            $Script:TotalPoolWaste = 0.00
            $Script:TotalUnassignedCount = 0
            $UnassignedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            
            # Fetch local tenant subbed SKUs and compute unassigned units
            try {
                $SubSkus = Get-MgSubscribedSku -All
                $Script:TenantSubSkus = $SubSkus
                foreach ($sku in $SubSkus) {
                    $GuidStr = $sku.SkuId.ToString().ToLower().Trim()
                    $PartNumber = if ($sku.SkuPartNumber) { $sku.SkuPartNumber.ToString().Trim() } elseif ($sku.SkuPartName) { $sku.SkuPartName.ToString().Trim() } else { "" }
                    
                    # Resolve SKU Name using GUID mapping, then PartNumber string mapping, then Global Cache, then formatted PartNumber
                    $SkuName = if ($Script:SkuMap.ContainsKey($GuidStr)) { 
                        $Script:SkuMap[$GuidStr] 
                    } elseif ($PartNumber -and $Script:SkuMap.ContainsKey($PartNumber.ToLower())) { 
                        $Script:SkuMap[$PartNumber.ToLower()] 
                    } elseif ($PartNumber -and $Script:SkuCacheByStringId.ContainsKey($PartNumber.ToUpper())) {
                        $Script:SkuCacheByStringId[$PartNumber.ToUpper()]
                    } elseif ($PartNumber) {
                        $clean = ($PartNumber -replace '_', ' ').Trim()
                        (Get-Culture).TextInfo.ToTitleCase($clean.ToLower())
                    } else { 
                        "Unknown SKU ($GuidStr)" 
                    }
                    
                    # Register into SkuMap for both GUID and PartNumber
                    $Script:SkuMap[$GuidStr] = $SkuName
                    if ($PartNumber) {
                        $Script:SkuMap[$PartNumber.ToLower()] = $SkuName
                    }
                    
                    $ActiveUnits = if ($sku.PrepaidUnits -and $sku.PrepaidUnits.Enabled) { $sku.PrepaidUnits.Enabled } else { 0 }
                    $Unassigned = $ActiveUnits - $sku.ConsumedUnits
                    $UnitPrice = Get-LicenseMonthlyPrice $SkuName
                    
                    # If this SKU is missing from the CSV database, estimate and append it
                    if ($null -ne $Script:SkuPrices -and -not $Script:SkuPrices.ContainsKey($GuidStr)) {
                        if ($SkuName -notlike "Unknown SKU*" -and -not [string]::IsNullOrWhiteSpace($SkuName)) {
                            try {
                                Add-Content -Path $PricesFile -Value "`n$SkuName,$UnitPrice,$GuidStr" -Encoding utf8
                                $Script:SkuPrices[$GuidStr] = $UnitPrice
                                $Script:SkuPrices[$SkuName] = $UnitPrice
                                Log-ToTerminal "Discovered tenant SKU '$SkuName' ($GuidStr). Added to LicensePrices.csv with estimate £{0:N2}." -f $UnitPrice
                            } catch {}
                        }
                    }
                    
                    if ($null -ne $Script:SkuPrices -and $Script:SkuPrices.ContainsKey($GuidStr)) {
                        $UnitPrice = $Script:SkuPrices[$GuidStr]
                    }
                    
                    # Skip free licenses (price == 0) or massive viral/trial plans (>= 10,000 active units) unless it contains "Teams"
                    if (($UnitPrice -eq 0.00 -or $ActiveUnits -ge 10000) -and $SkuName -notlike "*Teams*") {
                        continue
                    }
                    
                    if ($Unassigned -gt 0) {
                        $Script:TotalPoolWaste += ($Unassigned * $UnitPrice)
                        $Script:TotalUnassignedCount += $Unassigned
                    }
                    
                    $GreenBrush = $Window.Resources['BrushOk']
                    $RedBrush = $Window.Resources['BrushErr']
                    $GrayBrush = $Window.Resources['BrushTextMuted']
                    
                    $UnassignedTone = if ($Unassigned -gt 0) { "Ok" } elseif ($Unassigned -lt 0) { "Err" } else { "Muted" }
                    $WastedTone = if ($Unassigned -gt 0) { "Err" } elseif ($Unassigned -lt 0) { "Ok" } else { "Muted" }
                    $UnassignedBrush = if ($Unassigned -gt 0) { $GreenBrush } elseif ($Unassigned -lt 0) { $RedBrush } else { $GrayBrush }
                    $WastedBrush = if ($Unassigned -gt 0) { $RedBrush } elseif ($Unassigned -lt 0) { $GreenBrush } else { $GrayBrush }
                    
                    $UnassignedList.Add([PSCustomObject]@{
                        SkuPartName      = $SkuName
                        SkuId            = $GuidStr
                        ActiveUnits      = $ActiveUnits
                        ConsumedUnits    = $sku.ConsumedUnits
                        UnassignedUnits  = $Unassigned
                        UnassignedTone   = $UnassignedTone
                        WastedTone       = $WastedTone
                        UnassignedBrush  = $UnassignedBrush
                        WastedBrush      = $WastedBrush
                        MonthlyPrice     = $UnitPrice
                        MonthlyPriceText = "£{0:N2}" -f $UnitPrice
                        WastedCost       = ($Unassigned * $UnitPrice)
                        WastedCostText   = "£{0:N2}" -f ($Unassigned * $UnitPrice)
                        AssignedText     = "$($sku.ConsumedUnits) of $ActiveUnits"
                    })
                }
                $Script:UnassignedLicenses = $UnassignedList
            } catch {
                Log-ToTerminal "Failed to fetch local tenant SKUs. Using catalog reference instead." "Warning"
                Update-UI
            }
            
            Log-ToTerminal "Retrieving directory users (Licensed, Unlicensed, Guests, Enabled & Disabled)..." "Info"
            Update-UI
            
            $AllFetchedUsers = $null
            $HasSignInActivity = $true

            try {
                # Attempt 1: Query all users with SignInActivity (requires Entra ID P1/P2)
                $UserProperties = @('Id', 'DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AssignedPlans', 'SignInActivity', 'AccountEnabled', 'UserType', 'Mail')
                $AllFetchedUsers = Get-MgUser -All -Property $UserProperties -ErrorAction Stop
            } catch {
                if ($_ -match "Authentication_RequestFromNonPremiumTenantOrB2CTenant" -or $_ -match "premium license" -or $_ -match "403" -or $_ -match "SignInActivity") {
                    Log-ToTerminal "Notice: Tenant does not have Entra ID P1/P2 Premium license. Proceeding without sign-in dates..." "Warning"
                    Update-UI
                    $HasSignInActivity = $false
                    try {
                        # Attempt 2: Query without SignInActivity
                        $UserPropertiesBasic = @('Id', 'DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AssignedPlans', 'AccountEnabled', 'UserType', 'Mail')
                        $AllFetchedUsers = Get-MgUser -All -Property $UserPropertiesBasic -ErrorAction Stop
                    } catch {
                        throw $_
                    }
                } else {
                    Log-ToTerminal "Advanced user query failed ($($_)). Retrying standard query..." "Warning"
                    Update-UI
                    $HasSignInActivity = $false
                    try {
                        $UserPropertiesBasic = @('Id', 'DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AssignedPlans', 'AccountEnabled', 'UserType', 'Mail')
                        $AllFetchedUsers = Get-MgUser -All -Property $UserPropertiesBasic -ErrorAction Stop
                    } catch {
                        throw $_
                    }
                }
            }
            
            $UserCount = if ($AllFetchedUsers) { $AllFetchedUsers.Count } else { 0 }
            Log-ToTerminal "Fetched $UserCount directory users from Graph. Resolving licenses, user types, and telemetry..." "Info"
            Update-UI
            
            $Users = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($u in $AllFetchedUsers) {
                $Lics = @()
                if ($u.AssignedLicenses -and $u.AssignedLicenses.Count -gt 0) {
                    $Lics = foreach ($l in $u.AssignedLicenses) {
                        Resolve-SkuName $l.SkuId $u.AssignedPlans
                    }
                }
                $LicString = if ($Lics.Count -gt 0) { $Lics -join ", " } else { "None (Unlicensed)" }
                $IsLicensed = ($Lics.Count -gt 0)

                # Determine UserType (Member or Guest)
                $UserType = if ($u.UserType) {
                    $u.UserType
                } elseif ($u.UserPrincipalName -like "*#EXT#*" -or ($u.Mail -and $u.Mail -like "*#EXT#*")) {
                    "Guest"
                } else {
                    "Member"
                }
                $IsGuest = ($UserType -eq "Guest")
                
                $LastSignIn = if ($HasSignInActivity -and $u.SignInActivity) {
                    if ($u.SignInActivity.LastSuccessfulSignInDateTime) {
                        $u.SignInActivity.LastSuccessfulSignInDateTime
                    } elseif ($u.SignInActivity.LastSignInDateTime) {
                        $u.SignInActivity.LastSignInDateTime
                    } else {
                        $null
                    }
                } else { $null }

                if (-not $LastSignIn) {
                    if (-not $HasSignInActivity) {
                        $LastSignIn = "Requires Entra ID P1/P2"
                    } else {
                        $LastSignIn = "No interactive sign-in recorded"
                    }
                }
                
                $Users.Add([PSCustomObject]@{
                    DisplayName       = $u.DisplayName
                    UserPrincipalName = $u.UserPrincipalName
                    AssignedLicenses  = $LicString
                    IsLicensed        = $IsLicensed
                    UserType          = $UserType
                    IsGuest           = $IsGuest
                    LastSignInDate    = $LastSignIn
                    AccountEnabled    = $u.AccountEnabled
                })
            }
            
            Process-UserData $Users
            Log-ToTerminal "Report generation complete. Successfully processed $($Users.Count) tenant users (members, guests, active, and disabled)." "Success"
        } catch {
            Log-ToTerminal "Audit run failed: $_" "Error"
            [System.Windows.MessageBox]::Show("Audit Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    
    if ([string]::IsNullOrEmpty($Script:UserRole)) {
        $Script:UserRole = "Sales"
    }
    Configure-RoleUI
    
    $wpf_btnRunScript.IsEnabled = $true
    $wpf_btnRunScript.Content = "Refresh Report"
    Update-UI
    Show-Page "Dashboard"
}

function Parse-DateString($DateStr) {
    if ($null -eq $DateStr) { return $null }
    if ($DateStr -is [System.DateTime]) { return $DateStr }
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $null }
    if ($DateStr -like "*no interactive*" -or $DateStr -like "*recorded*" -or $DateStr -like "*Requires Entra ID*") { return $null }
    
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
    if ($null -eq $DateStr) { return [double]::PositiveInfinity }
    if ($DateStr -like "*Requires Entra ID*" -or $DateStr -like "*No P1/P2*") { return -1 }
    
    $Date = Parse-DateString $DateStr
    if ($null -eq $Date) { return [double]::PositiveInfinity }
    
    $Diff = (Get-Date) - $Date
    if ($Diff.TotalDays -lt 0) { return 0 }
    return [Math]::Floor($Diff.TotalDays)
}

function Format-LastSignIn($DateStr) {
    if ($null -eq $DateStr) { return "Never" }
    if ($DateStr -like "*Requires Entra ID*" -or $DateStr -like "*No P1/P2*") { return "Requires Entra ID P1/P2" }
    
    $Date = Parse-DateString $DateStr
    if ($null -eq $Date) { return "Never" }
    return $Date.ToString("dd/MM/yyyy HH:mm")
}

# 8. Charting Layout Engines (WPF Native Shapes)
function UpdateBarChart($Panel, $DataMap, $ColorKeyOrHex) {
    $Panel.Children.Clear()
    if ($null -eq $DataMap -or $DataMap.Count -eq 0) { return }
    
    $MaxVal = 1
    foreach ($val in $DataMap.Values) {
        if ($val -gt $MaxVal) { $MaxVal = $val }
    }
    
    foreach ($Key in $DataMap.Keys) {
        $Val = $DataMap[$Key]
        
        $Grid = [System.Windows.Controls.Grid]::new()
        $Grid.Margin = "0,6,0,6"
        
        $Col1 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col1.Width = [System.Windows.GridLength]::new(180)
        $Col2 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $Col3 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col3.Width = [System.Windows.GridLength]::new(40)
        
        $Grid.ColumnDefinitions.Add($Col1)
        $Grid.ColumnDefinitions.Add($Col2)
        $Grid.ColumnDefinitions.Add($Col3)
        
        $Label = [System.Windows.Controls.TextBlock]::new()
        $Label.Text = $Key
        $Label.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextMuted')
        $Label.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $Label.FontSize = 11
        $Label.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
        [System.Windows.Controls.Grid]::SetColumn($Label, 0)
        $Grid.Children.Add($Label)
        
        $BarOuter = [System.Windows.Controls.Border]::new()
        $BarOuter.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushInputBg')
        $BarOuter.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushCardBorder')
        $BarOuter.BorderThickness = 1
        $BarOuter.CornerRadius = 3
        $BarOuter.Height = 14
        $BarOuter.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        $BarOuter.Width = 200
        [System.Windows.Controls.Grid]::SetColumn($BarOuter, 1)
        
        $BarWidth = [Math]::Round(($Val / $MaxVal) * 198)
        if ($BarWidth -lt 2) { $BarWidth = 2 }
        $BarInner = [System.Windows.Shapes.Rectangle]::new()
        if ($ColorKeyOrHex -like 'Brush*') {
            $BarInner.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $ColorKeyOrHex)
        } else {
            $BarInner.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($ColorKeyOrHex)
        }
        $BarInner.Width = $BarWidth
        $BarInner.Height = 12
        $BarInner.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        $BarInner.Margin = "1,0,0,0"
        
        $BarOuter.Child = $BarInner
        $Grid.Children.Add($BarOuter)
        
        $ValLabel = [System.Windows.Controls.TextBlock]::new()
        $ValLabel.Text = $Val.ToString()
        $ValLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
        $ValLabel.FontWeight = [System.Windows.FontWeights]::Bold
        $ValLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $ValLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        $ValLabel.FontSize = 11
        [System.Windows.Controls.Grid]::SetColumn($ValLabel, 2)
        $Grid.Children.Add($ValLabel)
        
        $Panel.Children.Add($Grid)
    }
}

function AddActivityBarChartRow($Panel, $LabelText, $Val, $Total, $ColorKeyOrHex) {
    if ($Total -eq 0) { $Total = 1 }
    
    $Grid = [System.Windows.Controls.Grid]::new()
    $Grid.Margin = "0,6,0,6"
    
    $Col1 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col1.Width = [System.Windows.GridLength]::new(180)
    $Col2 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $Col3 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col3.Width = [System.Windows.GridLength]::new(40)
    
    $Grid.ColumnDefinitions.Add($Col1)
    $Grid.ColumnDefinitions.Add($Col2)
    $Grid.ColumnDefinitions.Add($Col3)
    
    $Label = [System.Windows.Controls.TextBlock]::new()
    $Label.Text = $LabelText
    $Label.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextMuted')
    $Label.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $Label.FontSize = 11
    [System.Windows.Controls.Grid]::SetColumn($Label, 0)
    $Grid.Children.Add($Label)
    
    $BarOuter = [System.Windows.Controls.Border]::new()
    $BarOuter.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushInputBg')
    $BarOuter.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushCardBorder')
    $BarOuter.BorderThickness = 1
    $BarOuter.CornerRadius = 3
    $BarOuter.Height = 14
    $BarOuter.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    $BarOuter.Width = 200
    [System.Windows.Controls.Grid]::SetColumn($BarOuter, 1)
    
    $BarWidth = [Math]::Round(($Val / $Total) * 198)
    if ($BarWidth -lt 2) { $BarWidth = 2 }
    $BarInner = [System.Windows.Shapes.Rectangle]::new()
    if ($ColorKeyOrHex -like 'Brush*') {
        $BarInner.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $ColorKeyOrHex)
    } else {
        $BarInner.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($ColorKeyOrHex)
    }
    $BarInner.Width = $BarWidth
    $BarInner.Height = 12
    $BarInner.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    $BarInner.Margin = "1,0,0,0"
    
    $BarOuter.Child = $BarInner
    $Grid.Children.Add($BarOuter)
    
    $ValLabel = [System.Windows.Controls.TextBlock]::new()
    $ValLabel.Text = $Val.ToString()
    $ValLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
    $ValLabel.FontWeight = [System.Windows.FontWeights]::Bold
    $ValLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $ValLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $ValLabel.FontSize = 11
    [System.Windows.Controls.Grid]::SetColumn($ValLabel, 2)
    $Grid.Children.Add($ValLabel)
    
    $Panel.Children.Add($Grid)
}

# 9. Dashboard Processors
function Process-UserData($Users) {
    $Script:CurrentData = $Users
    
    $Total = $Users.Count
    $Active = 0
    $Inactive90d = 0
    $Inactive1yr = 0
    $Never = 0
    $NoP1P2Count = 0
    
    $LicenseMap = @{}
    $ActivityMap = [ordered]@{}
    
    $HasNoP1P2 = $false
    foreach ($user in $Users) {
        if ($user.LastSignInDate -like "*Requires Entra ID*" -or $user.LastSignInDate -like "*No P1/P2*") {
            $HasNoP1P2 = $true
            break
        }
    }
    
    if ($HasNoP1P2) {
        $ActivityMap["Requires Entra ID P1/P2"] = 0
    }
    $ActivityMap["Active (<=30d)"]    = 0
    $ActivityMap["Inactive (30-90d)"]  = 0
    $ActivityMap["Inactive (90-365d)"] = 0
    $ActivityMap["Inactive (>1yr)"]    = 0
    $ActivityMap["Never Logged In"]    = 0

    $DataGridRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($user in $Users) {
        $Days = Get-DaysSince $user.LastSignInDate
        
        $StatusText = ""
        $Category = ""
        if ($Days -eq -1) {
            $StatusText = "No P1/P2 License"
            $Category = "Requires Entra ID P1/P2"
            $NoP1P2Count++
        } elseif ($Days -eq [double]::PositiveInfinity) {
            $StatusText = "Never Logged In"
            $Category = "Never Logged In"
            $Never++
        } elseif ($Days -le 30) {
            $StatusText = "Active ($Days`d ago)"
            $Category = "Active (<=30d)"
            $Active++
        } elseif ($Days -le 90) {
            $StatusText = "Warning ($Days`d ago)"
            $Category = "Inactive (30-90d)"
        } elseif ($Days -le 365) {
            $StatusText = "Inactive ($Days`d/90+ ago)"
            $Category = "Inactive (90-365d)"
            $Inactive90d++
        } else {
            $StatusText = "Critical ($Days`d/1yr+ ago)"
            $Category = "Inactive (>1yr)"
            $Inactive1yr++
            $Inactive90d++
        }
        
        if ($ActivityMap.Contains($Category)) {
            $ActivityMap[$Category]++
        } else {
            $ActivityMap[$Category] = 1
        }
        
        if ($user.AssignedLicenses -and $user.AssignedLicenses -ne "None (Unlicensed)") {
            $user.AssignedLicenses.Split(',') | ForEach-Object {
                $lic = $_.Trim()
                if ($lic -and $lic -ne "None (Unlicensed)") {
                    $LicenseMap[$lic]++
                }
            }
        }
        
        $Username = ($user.UserPrincipalName -split '@')[0].ToLower()
        $UpnLower = ([string]$user.UserPrincipalName).ToLower()
        $DisplayNameLower = ([string]$user.DisplayName).ToLower()
        $VerificationText = "-"
        foreach ($Pattern in $Script:ExcludeAccountPatterns) {
            if ($UpnLower -like $Pattern -or $Username -like $Pattern -or $DisplayNameLower -like $Pattern) {
                $VerificationText = "MSP Verification Account"
                break
            }
        }

        $Enabled = $true
        if ($null -ne $user.AccountEnabled) {
            $Enabled = $user.AccountEnabled
        }
        $AccountStatusText = if ($Enabled) { "Enabled" } else { "Disabled (Blocked)" }

        $WastedCost = 0.00
        $MonthlySavings = 0.00
        if ($user.AssignedLicenses -and $user.AssignedLicenses -ne "None (Unlicensed)") {
            if ($Days -gt 180 -and $Days -ne [double]::PositiveInfinity -and $Days -ne -1) {
                $InactiveDays = $Days
                $Months = $InactiveDays / 30
                $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
                $WastedCost = $MonthlyPrice * $Months
                $MonthlySavings = $MonthlyPrice
            } elseif ($Days -eq [double]::PositiveInfinity) {
                $InactiveDays = 365
                $Months = $InactiveDays / 30
                $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
                $WastedCost = $MonthlyPrice * $Months
                $MonthlySavings = $MonthlyPrice
            }
        }
        $WastedCostText = "£{0:N2}" -f $WastedCost
        $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings
        
        # Downgrade / Reclamation Recommendation Engine
        $Recommendation = "-"
        if (-not $Enabled -and $user.AssignedLicenses -and $user.AssignedLicenses -ne "None (Unlicensed)") {
            $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
            if ($MonthlyPrice -gt 0) {
                $Recommendation = "Reclaim: Account disabled with active license (Save £{0:N2}/mo)" -f $MonthlyPrice
                $MonthlySavings = $MonthlyPrice
                $WastedCost = if ($Days -gt 0 -and $Days -ne [double]::PositiveInfinity -and $Days -ne -1) { $MonthlyPrice * ($Days / 30) } else { $MonthlyPrice * 6 }
                $WastedCostText = "£{0:N2}" -f $WastedCost
                $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings
            }
        } elseif ($user.AssignedLicenses -and $user.AssignedLicenses -ne "None (Unlicensed)") {
            $LicsArray = $user.AssignedLicenses.Split(',') | ForEach-Object { $_.Trim() }
            $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
            
            if ($Days -eq [double]::PositiveInfinity -or ($Days -gt 180 -and $Days -ne -1)) {
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
                $ExchangeP1Price = if ($Script:SkuPrices.ContainsKey("Exchange Online (Plan 1)")) { $Script:SkuPrices["Exchange Online (Plan 1)"] } else { 3.30 }
                if ($LicsArray -contains "Exchange Online (Plan 1)" -or $LicsArray -contains "EXCHANGESTANDARD") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3 -or $HasBusinessStandard -or $HasBusinessBasic) {
                        $RedundantPlans.Add("Exchange P1")
                        $RedundantCost += $ExchangeP1Price
                    }
                }
                
                # 2. Exchange Online (Plan 2) redundant check (BP mailboxes upgrade to 100GB as of July 1, 2026!)
                $ExchangeP2Price = if ($Script:SkuPrices.ContainsKey("Exchange Online (Plan 2)")) { $Script:SkuPrices["Exchange Online (Plan 2)"] } else { 6.60 }
                if ($LicsArray -contains "Exchange Online (Plan 2)" -or $LicsArray -contains "EXCHANGEENTERPRISE") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                        $RedundantPlans.Add("Exchange P2")
                        $RedundantCost += $ExchangeP2Price
                    }
                }
                
                # 3. Microsoft Entra ID P1 / AAD Premium redundant check
                $EntraP1Price = if ($Script:SkuPrices.ContainsKey("Microsoft Entra ID P1")) { $Script:SkuPrices["Microsoft Entra ID P1"] } else { 4.90 }
                if ($LicsArray -contains "Microsoft Entra ID P1" -or $LicsArray -contains "AAD_PREMIUM") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                        $RedundantPlans.Add("Entra ID P1")
                        $RedundantCost += $EntraP1Price
                    }
                }
                
                # 4. Microsoft Intune redundant check
                $IntunePrice = if ($Script:SkuPrices.ContainsKey("Microsoft Intune")) { $Script:SkuPrices["Microsoft Intune"] } else { 6.60 }
                if ($LicsArray -contains "Microsoft Intune" -or $LicsArray -contains "INTUNE_A") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3) {
                        $RedundantPlans.Add("Intune")
                        $RedundantCost += $IntunePrice
                    }
                }

                # 5. Power BI Pro redundant check
                $PowerBIPrice = if ($Script:SkuPrices.ContainsKey("Power BI Pro")) { $Script:SkuPrices["Power BI Pro"] } else { 8.20 }
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

        $UserTypeVal = if ($user.UserType) { $user.UserType } elseif ($user.UserPrincipalName -like "*#EXT#*") { "Guest" } else { "Member" }
        $IsGuestVal = ($UserTypeVal -eq "Guest")
        $IsLicensedVal = ($user.AssignedLicenses -and $user.AssignedLicenses -ne "None (Unlicensed)")

        $DataGridRows.Add([PSCustomObject]@{
            DisplayName        = $user.DisplayName
            UserPrincipalName  = $user.UserPrincipalName
            UserType           = $UserTypeVal
            IsGuest            = $IsGuestVal
            IsLicensed         = $IsLicensedVal
            AccountEnabled     = $Enabled
            AssignedLicenses   = $user.AssignedLicenses
            LastSignInDate     = (Format-LastSignIn $user.LastSignInDate)
            StatusText         = $StatusText
            StatusCategory     = $Category
            DaysSince          = $Days
            VerificationText   = $VerificationText
            WastedCost         = $WastedCost
            WastedCostText     = $WastedCostText
            MonthlySavings     = $MonthlySavings
            MonthlySavingsText = $MonthlySavingsText
            Recommendation     = $Recommendation
            AccountStatusText  = $AccountStatusText
        })
    }
    
    $Script:GridRows = $DataGridRows
    
    # Calculate Total Wasted Cost and Monthly Savings
    $TotalSavings = 0.00
    $TotalMonthlySavings = 0.00
    foreach ($row in $DataGridRows) {
        $TotalSavings += $row.WastedCost
        $TotalMonthlySavings += $row.MonthlySavings
    }
    
    # Update KPI UI Text blocks
    $wpf_kpiTotal.Text = $Total.ToString()
    $wpf_kpiActive.Text = $Active.ToString()
    $wpf_kpiInactive90d.Text = $Inactive90d.ToString()
    $wpf_kpiInactive1yr.Text = $Inactive1yr.ToString()
    $wpf_kpiSavings.Text = "£{0:N2}" -f $TotalSavings
    $wpf_kpiMonthlySavings.Text = "£{0:N2}/mo" -f $TotalMonthlySavings
    
    $PoolWasteVal = if ($null -ne $Script:TotalPoolWaste) { $Script:TotalPoolWaste } else { 0.00 }
    $wpf_kpiPoolWaste.Text = "£{0:N2}" -f $PoolWasteVal
    
    $PoolWasteCountVal = if ($null -ne $Script:TotalUnassignedCount) { $Script:TotalUnassignedCount } else { 0 }
    $wpf_kpiPoolWasteCount.Text = "$PoolWasteCountVal unused licenses"
    
    # Draw License Bar Chart
    if ($wpf_panelLicenseChart) { UpdateBarChart $wpf_panelLicenseChart $LicenseMap "BrushAccent" }
    
    # Draw Activity Bar Chart with distinct colors
    $wpf_panelActivityChart.Children.Clear()
    foreach ($Key in $ActivityMap.Keys) {
        $Val = $ActivityMap[$Key]
        $Color = $Script:ActivityColors[$Key]
        AddActivityBarChartRow $wpf_panelActivityChart $Key $Val $Total $Color
    }
    
    if ($wpf_gridUnassigned) {
        $wpf_gridUnassigned.ItemsSource = $null
        $wpf_gridUnassigned.ItemsSource = $Script:UnassignedLicenses
    }
    
    Filter-DataGrid
}

function Filter-DataGrid {
    if ($null -eq $Script:GridRows) { return }
    
    $Filtered = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($row in $Script:GridRows) {
        # 1. Search Box Match (Matches DisplayName, UPN, AssignedLicenses, or UserType)
        if (-not [string]::IsNullOrWhiteSpace($Script:SearchQuery)) {
            $Q = $Script:SearchQuery.ToLower()
            $MatchName = $row.DisplayName -and $row.DisplayName.ToLower().Contains($Q)
            $MatchUPN = $row.UserPrincipalName -and $row.UserPrincipalName.ToLower().Contains($Q)
            $MatchLic = $row.AssignedLicenses -and $row.AssignedLicenses.ToLower().Contains($Q)
            $MatchType = $row.UserType -and $row.UserType.ToLower().Contains($Q)
            if (-not ($MatchName -or $MatchUPN -or $MatchLic -or $MatchType)) {
                continue
            }
        }
        
        # 2. Preset Filter Match
        $PresetMatch = switch ($Script:ActivePreset) {
            "all"              { $true }
            "active"           { $row.StatusCategory -eq "Active (<=30d)" }
            "licensed_guests"  { $row.IsLicensed -or $row.IsGuest }
            "licensed"         { $row.IsLicensed }
            "guests"           { $row.IsGuest }
            "unlicensed"       { -not $row.IsLicensed }
            "disabled"         { -not $row.AccountEnabled }
            "inactive90d"      { $row.DaysSince -ge 90 -and $row.DaysSince -ne -1 }
            "inactive1yr"      { $row.StatusCategory -eq "Inactive (>1yr)" }
            "never"            { $row.StatusCategory -eq "Never Logged In" }
            "custom"           { $true }
            default            { $true }
        }
        if (-not $PresetMatch) { continue }
        
        # 3. Additive Sub-Filter Refinements
        # License: all | licensed | unlicensed
        if ($Script:FilterLicense -eq "licensed" -and -not $row.IsLicensed) { continue }
        if ($Script:FilterLicense -eq "unlicensed" -and $row.IsLicensed) { continue }

        # User Type: all | members | guests | licensed_guests
        if ($Script:FilterUserType -eq "members" -and $row.IsGuest) { continue }
        if ($Script:FilterUserType -eq "guests" -and -not $row.IsGuest) { continue }
        if ($Script:FilterUserType -eq "licensed_guests" -and -not ($row.IsLicensed -or $row.IsGuest)) { continue }

        # Account Status: all | enabled | disabled
        if ($Script:FilterAccount -eq "enabled" -and -not $row.AccountEnabled) { continue }
        if ($Script:FilterAccount -eq "disabled" -and $row.AccountEnabled) { continue }

        # Sign-In Activity: all | active | inactive90d | never
        if ($Script:FilterActivity -eq "active" -and $row.StatusCategory -ne "Active (<=30d)") { continue }
        if ($Script:FilterActivity -eq "inactive90d" -and -not ($row.DaysSince -ge 90 -and $row.DaysSince -ne -1)) { continue }
        if ($Script:FilterActivity -eq "never" -and $row.StatusCategory -ne "Never Logged In") { continue }

        $Filtered.Add($row)
    }
    
    $wpf_gridUsers.ItemsSource = $null
    $wpf_gridUsers.ItemsSource = $Filtered

    # Update live count indicator
    $TotalCount = if ($Script:GridRows) { $Script:GridRows.Count } else { 0 }
    if ($wpf_txtFilterCount) {
        $wpf_txtFilterCount.Text = "Showing $($Filtered.Count) of $TotalCount users"
    }
}

# 10. Navigations & Tab Switchers
function Show-Page($PageName) {
    if ($wpf_pageConnection) { $wpf_pageConnection.Visibility = [System.Windows.Visibility]::Visible }
    if ($wpf_pageDashboard) { $wpf_pageDashboard.Visibility = [System.Windows.Visibility]::Visible }
    if ($wpf_pageDirectory) { $wpf_pageDirectory.Visibility = [System.Windows.Visibility]::Visible }

    if ($null -eq $wpf_tabMain) { return }
    switch ($PageName) {
        "Connection"  { if ($wpf_tabConnection)  { $wpf_tabMain.SelectedItem = $wpf_tabConnection } }
        "Dashboard"   { if ($wpf_tabOverview)    { $wpf_tabMain.SelectedItem = $wpf_tabOverview } }
        "Licences"    { if ($wpf_tabLicences)    { $wpf_tabMain.SelectedItem = $wpf_tabLicences } }
        "Directory"   { if ($wpf_tabDirectory)   { $wpf_tabMain.SelectedItem = $wpf_tabDirectory } }
        "ActivityLog" { if ($wpf_tabActivityLog) { $wpf_tabMain.SelectedItem = $wpf_tabActivityLog } }
    }
}

function Show-ConnectionMethod($MethodName) {
    if ($null -eq $wpf_panelMethodM365 -or $null -eq $wpf_panelMethodApp -or $null -eq $wpf_panelMethodCSV) { return }
    $wpf_panelMethodM365.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelMethodApp.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelMethodCSV.Visibility = [System.Windows.Visibility]::Collapsed

    foreach ($btn in @($wpf_btnMethodM365, $wpf_btnMethodApp, $wpf_btnMethodCSV)) {
        if ($btn) {
            $btn.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushSecBtnBg')
            $btn.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushTextMuted')
            $btn.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
            $btn.BorderThickness = 1
        }
    }

    if ($MethodName -eq "M365") {
        $wpf_panelMethodM365.Visibility = [System.Windows.Visibility]::Visible
        if ($wpf_btnMethodM365) {
            $wpf_btnMethodM365.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccent')
            $wpf_btnMethodM365.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushAccentFg')
            $wpf_btnMethodM365.BorderThickness = 0
        }
    } elseif ($MethodName -eq "App") {
        $wpf_panelMethodApp.Visibility = [System.Windows.Visibility]::Visible
        if ($wpf_btnMethodApp) {
            $wpf_btnMethodApp.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccent')
            $wpf_btnMethodApp.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushAccentFg')
            $wpf_btnMethodApp.BorderThickness = 0
        }
    } elseif ($MethodName -eq "CSV") {
        $wpf_panelMethodCSV.Visibility = [System.Windows.Visibility]::Visible
        if ($wpf_btnMethodCSV) {
            $wpf_btnMethodCSV.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccent')
            $wpf_btnMethodCSV.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushAccentFg')
            $wpf_btnMethodCSV.BorderThickness = 0
        }
    }
}

function Select-Role($Role) {
    $Script:UserRole = $Role

    if ($Role -eq "ServiceDesk") {
        $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccentTint')
        $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushAccent')

        $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushCardBg')
        $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
    } else {
        $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushCardBg')
        $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')

        $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccentTint')
        $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushAccent')
    }

    Configure-RoleUI

    # Reveal Step 2 connection methods panel
    $wpf_panelConnectionMethods.Visibility = [System.Windows.Visibility]::Visible
    Update-UI
}

$wpf_btnRoleSelectServiceDesk.Add_Click({ Select-Role "ServiceDesk" })
$wpf_btnRoleSelectSales.Add_Click({ Select-Role "Sales" })
if ($wpf_btnFooterActivity) {
    $wpf_btnFooterActivity.Add_Click({ Show-Page "ActivityLog" })
}

# 11. Event Handlers & Connection Logic
# Connect: M365 Interactive Login
$wpf_btnConnectM365.Add_Click({
    $TenantId = $wpf_txtUserTenant.Text.Trim()
    
    Log-ToTerminal "Initializing Microsoft Graph interactive login flow..." "Info"
    Log-ToTerminal "Connect-MgGraph -Scopes 'User.Read.All', 'AuditLog.Read.All', 'Directory.Read.All', 'Organization.Read.All'" "Command"
    Update-UI
    
    try {
        # Minimize window to allow browser authentication to take focus
        $Window.WindowState = [System.Windows.WindowState]::Minimized
        Update-UI
        Start-Sleep -Milliseconds 200
        
        # Clear Graph cache to ensure a clean start
        Clear-MgGraphCache
        
        $Params = @{
            Scopes = @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Organization.Read.All", "LicenseAssignment.Read.All")
            ContextScope = "Process"
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $Params.TenantId = $TenantId }
        
        Connect-MgGraph @Params | Out-Null
        
        # Restore window state and focus
        $Window.WindowState = [System.Windows.WindowState]::Normal
        $Window.Activate() | Out-Null
        
        $Context = Get-MgContext
        if ($null -eq $Context) { throw "Graph connection context could not be established." }
        
        $Script:SessionMode = "Real"
        $wpf_txtStatusLabel.Text = "Connected to Graph"
        $wpf_txtStatusLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
        if ($wpf_bdgSessionStatus) {
            $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushOkTint')
            $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushOk')
        }
        
        $TenantName = "UnknownTenant"
        try {
            $OrgInfo = Get-MgOrganization | Select-Object -First 1 -Property DisplayName
            $TenantName = $OrgInfo.DisplayName
        } catch {}
        
        $wpf_txtSessionTenant.Text = $TenantName
        $wpf_txtSessionUser.Text = $Context.Account
        $wpf_txtSessionMode.Text = "DIRECT GRAPH API"
        $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        
        $wpf_btnRunScript.IsEnabled = $true
        $wpf_btnRunScript.Opacity = 1.0
        
        Start-SessionTimer
        
        Log-ToTerminal "Successfully connected to Microsoft Graph." "Success"
        Log-ToTerminal "Account: $($Context.Account)" "Info"
        Log-ToTerminal "Tenant ID: $($Context.TenantId)" "Info"
        
        # Switch visibilities to connected session look
        $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
        $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
        
        Show-Page "Dashboard"
        Invoke-AuditReport
    } catch {
        # Restore window state in case of exception
        $Window.WindowState = [System.Windows.WindowState]::Normal
        $Window.Activate() | Out-Null
        Log-ToTerminal "Authentication failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Connection Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Connect: App Secret Token
$wpf_btnConnectApp.Add_Click({
    $TenantId = $wpf_txtAppTenant.Text.Trim()
    $ClientId = $wpf_txtAppClient.Text.Trim()
    $Secret = $wpf_txtAppSecret.Password.Trim()
    
    if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($Secret)) {
        [System.Windows.MessageBox]::Show("Please fill in the Tenant ID, Client ID, and Client Secret.", "Missing Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Log-ToTerminal "Connecting via Application token credentials..." "Info"
    Log-ToTerminal "Connect-MgGraph -TenantId '$TenantId' -ClientId '$ClientId' -ClientSecret '••••••••'" "Command"
    Update-UI
    
    try {
        # Clear Graph cache to ensure a clean start
        Clear-MgGraphCache
        
        $SecSecret = ConvertTo-SecureString $Secret -AsPlainText -Force
        $Credential = [PSCredential]::new($ClientId, $SecSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -ContextScope Process | Out-Null
        
        $Context = Get-MgContext
        if ($null -eq $Context) { throw "Graph connection context could not be established." }
        
        $Script:SessionMode = "AppRegistration"
        $wpf_txtStatusLabel.Text = "Connected (App Secret)"
        $wpf_txtStatusLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
        if ($wpf_bdgSessionStatus) {
            $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushOkTint')
            $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushOk')
        }
        
        $TenantName = "UnknownTenant"
        try {
            $OrgInfo = Get-MgOrganization | Select-Object -First 1 -Property DisplayName
            $TenantName = $OrgInfo.DisplayName
        } catch {}
        
        $wpf_txtSessionTenant.Text = $TenantName
        $wpf_txtSessionUser.Text = "AppRegistration ($ClientId)"
        $wpf_txtSessionMode.Text = "APP TOKEN REG"
        $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        
        $wpf_btnRunScript.IsEnabled = $true
        $wpf_btnRunScript.Opacity = 1.0
        
        Start-SessionTimer
        
        Log-ToTerminal "Successfully connected using Application Client Secret." "Success"
        
        # Switch visibilities to connected session look
        $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
        $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
        
        Show-Page "Dashboard"
        Invoke-AuditReport
    } catch {
        Log-ToTerminal "Connection failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Connection Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Connect: Browse and Import local CSV report outputs
$wpf_btnBrowseCSV.Add_Click({
    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Import PowerShell Tenant CSV Report"
    
    if ($openFileDialog.ShowDialog()) {
        $File = $openFileDialog.FileName
        Log-ToTerminal "Importing local CSV file: $($openFileDialog.SafeFileName)..." "Info"
        Update-UI
        try {
            $CsvData = Import-Csv -Path $File
            if ($CsvData.Count -eq 0) { throw "CSV file contains no records." }
            
            $Headers = $CsvData[0].PSObject.Properties.Name
            
            $DN_Prop = $Headers | Where-Object { $_ -eq "DisplayName" -or $_ -eq "Display Name" -or $_ -eq "Display_Name" } | Select-Object -First 1
            $UPN_Prop = $Headers | Where-Object { $_ -eq "UserPrincipalName" -or $_ -eq "User Principal Name" -or $_ -eq "User_Principal_Name" -or $_ -eq "UPN" } | Select-Object -First 1
            $Lic_Prop = $Headers | Where-Object { $_ -eq "AssignedLicenses" -or $_ -eq "Assigned Licenses" -or $_ -eq "Licenses" } | Select-Object -First 1
            $Date_Prop = $Headers | Where-Object { $_ -eq "LastSignInDate" -or $_ -eq "Last Sign-In Date" -or $_ -eq "Last_Sign-In_Date" -or $_ -eq "LastSignIn" } | Select-Object -First 1
            
            if (-not ($DN_Prop -and $UPN_Prop -and $Lic_Prop -and $Date_Prop)) {
                throw "Missing required headers. Columns must include Display Name, User Principal Name, Assigned Licenses, and Last Sign-In Date."
            }
            
            if ([string]::IsNullOrEmpty($Script:UserRole)) {
                $Script:UserRole = "Sales"
            }
            Configure-RoleUI
            
            # Extract license counts from imported CSV to show in the breakdown
            $CsvSkuCounts = @{}
            foreach ($row in $CsvData) {
                if ($row.$Lic_Prop) {
                    $row.$Lic_Prop.Split(',') | ForEach-Object {
                        $SkuName = $_.Trim()
                        if ($SkuName) {
                            $CsvSkuCounts[$SkuName] = ($CsvSkuCounts[$SkuName] + 1)
                        }
                    }
                }
            }
            
            $UnassignedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
            $PricesCsv = if (Test-Path $PricesFile) { Import-Csv $PricesFile } else { @() }
            
            foreach ($SkuName in $CsvSkuCounts.Keys) {
                $UnitPrice = Get-LicenseMonthlyPrice $SkuName
                
                # Resolve Sku ID GUID from LicensePrices.csv if available
                $ResolvedSkuId = "-"
                if ($PricesCsv) {
                    $MatchedRow = $PricesCsv | Where-Object { $_.Product -eq $SkuName -or $_.SkuId -eq $SkuName } | Select-Object -First 1
                    if ($MatchedRow) {
                        $ResolvedSkuId = $MatchedRow.SkuId
                    }
                }
                
                $Count = $CsvSkuCounts[$SkuName]
                $GrayBrush = $Window.Resources['BrushTextMuted']
                $UnassignedList.Add([PSCustomObject]@{
                    SkuPartName      = $SkuName
                    SkuId            = $ResolvedSkuId
                    ActiveUnits      = $Count
                    ConsumedUnits    = $Count
                    UnassignedUnits  = 0
                    UnassignedTone   = "Muted"
                    WastedTone       = "Muted"
                    UnassignedBrush  = $GrayBrush
                    WastedBrush      = $GrayBrush
                    MonthlyPrice     = $UnitPrice
                    MonthlyPriceText = "£{0:N2}" -f $UnitPrice
                    WastedCost       = 0.00
                    WastedCostText   = "£0.00"
                    AssignedText     = "$Count of $Count"
                })
            }
            
            $Script:UnassignedLicenses = $UnassignedList
            $Script:TotalPoolWaste = 0.00
            $Script:TotalUnassignedCount = 0
            
            $Script:SessionMode = "CSVImport"
            $wpf_txtStatusLabel.Text = "Loaded Local CSV"
            $wpf_txtStatusLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
            if ($wpf_bdgSessionStatus) {
                $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushOkTint')
                $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushOk')
            }
            
            $TenantName = "Local CSV"
            if ($openFileDialog.SafeFileName -match "LicensedUsers_([^_]+)_") {
                $TenantName = $Matches[1]
            }
            
            $wpf_txtSessionTenant.Text = $TenantName
            $wpf_txtSessionUser.Text = $openFileDialog.SafeFileName
            $wpf_txtSessionMode.Text = "CSV FILE IMPORT"
            $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
            
            $wpf_btnRunScript.IsEnabled = $false
            $wpf_btnRunScript.Opacity = 0.5
            
            $Users = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($row in $CsvData) {
                # Look for AccountEnabled in CSV
                $AccEnabled = $true
                $AccProp = $Headers | Where-Object { $_ -eq "AccountEnabled" -or $_ -eq "Account Status" -or $_ -eq "Account_Status" } | Select-Object -First 1
                if ($AccProp -and $row.$AccProp) {
                    if ($row.$AccProp -match "Blocked" -or $row.$AccProp -eq "False" -or $row.$AccProp -eq $false) {
                        $AccEnabled = $false
                    }
                }
                # Look for UserType in CSV
                $TypeProp = $Headers | Where-Object { $_ -eq "UserType" -or $_ -eq "User Type" -or $_ -eq "Account Type" -or $_ -eq "Type" } | Select-Object -First 1
                $UserType = if ($TypeProp -and $row.$TypeProp) { $row.$TypeProp } elseif ($row.$UPN_Prop -like "*#EXT#*") { "Guest" } else { "Member" }

                $Users.Add([PSCustomObject]@{
                    DisplayName       = $row.$DN_Prop
                    UserPrincipalName = $row.$UPN_Prop
                    AssignedLicenses  = $row.$Lic_Prop
                    UserType          = $UserType
                    LastSignInDate    = $row.$Date_Prop
                    AccountEnabled    = $AccEnabled
                })
            }
            
            Process-UserData $Users
            Log-ToTerminal "Successfully imported $($Users.Count) records from CSV." "Success"
            
            # Switch visibilities to connected session look
            $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
            $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
            
            Show-Page "Dashboard"
        } catch {
            Log-ToTerminal "CSV Import failed: $_" "Error"
            [System.Windows.MessageBox]::Show("CSV Import Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Connect: Demo Sandbox Mode
$wpf_btnDemoMode.Add_Click({
    $Script:SessionMode = "Demo"
    $wpf_txtStatusLabel.Text = "Active Sandbox Session"
    $wpf_txtStatusLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')
    if ($wpf_bdgSessionStatus) {
        $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushOkTint')
        $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushOk')
    }
    
    $wpf_txtSessionTenant.Text = "contoso.com (Demo)"
    $wpf_txtSessionUser.Text = "admin_demo@contoso.com"
    $wpf_txtSessionMode.Text = "DEMO STANDALONE"
    $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
    
    $wpf_btnRunScript.IsEnabled = $true
    $wpf_btnRunScript.Opacity = 1.0
    
    Start-SessionTimer
    
    Log-ToTerminal "Initializing local PowerShell sandbox variables..." "Info"
    Log-ToTerminal "Import-Module Microsoft.Graph -Force" "Command"
    Log-ToTerminal "Microsoft.Graph SDK modules loaded successfully." "Success"
    Log-ToTerminal "Connect-MgGraph -DemoSandboxMode" "Command"
    Log-ToTerminal "Connected in offline sandbox session environment." "Success"
    
    # Switch visibilities to connected session look
    $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
    
    Show-Page "Dashboard"
    Invoke-AuditReport
})

# Disconnect Session
$wpf_btnDisconnect.Add_Click({
    Stop-SessionTimer
    $Script:SessionMode = ""
    $Script:CurrentData = $null
    $Script:GridRows = $null
    $Script:UserRole = ""
    
    $wpf_txtStatusLabel.Text = "Offline Session"
    $wpf_txtStatusLabel.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextMuted')
    if ($wpf_bdgSessionStatus) {
        $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushSecBtnBg')
        $wpf_bdgSessionStatus.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushCardBorder')
    }
    
    $wpf_txtSessionTenant.Text = "N/A"
    $wpf_txtSessionUser.Text = "N/A"
    $wpf_txtSessionUptime.Text = "00:00:00"
    $wpf_txtSessionMode.Text = "N/A"
    $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_btnRunScript.IsEnabled = $false
    
    $wpf_txtAppTenant.Text = ""
    $wpf_txtAppClient.Text = ""
    $wpf_txtAppSecret.Password = ""
    $wpf_txtUserTenant.Text = ""
    
    $wpf_kpiTotal.Text = "0"
    $wpf_kpiActive.Text = "0"
    $wpf_kpiInactive90d.Text = "0"
    $wpf_kpiInactive1yr.Text = "0"
    
    if ($wpf_panelLicenseChart) { $wpf_panelLicenseChart.Children.Clear() }
    $wpf_panelActivityChart.Children.Clear()
    $wpf_gridUsers.ItemsSource = $null
    
    # Reset role cards and hide connection options panel
    $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Visible
    $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Collapsed
    
    $wpf_panelConnectionMethods.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushCardBg')
    $wpf_btnRoleSelectServiceDesk.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
    $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushCardBg')
    $wpf_btnRoleSelectSales.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
    
    Clear-MgGraphCache
    
    Log-ToTerminal "Disconnecting active Microsoft Graph session..." "Info"
    Log-ToTerminal "Disconnect-MgGraph" "Command"
    Log-ToTerminal "Session disconnected cleanly." "Success"
    Show-Page "Connection"
})

# Execute Audit script simulation / native queries
$wpf_btnRunScript.Add_Click({
    Invoke-AuditReport
})

# Directory Search Text Changes
$wpf_txtSearch.Add_TextChanged({
    $Script:SearchQuery = $wpf_txtSearch.Text.Trim()
    Filter-DataGrid
})

# Directory Filter Selection Changes
$Script:PresetButtons = @(
    $wpf_btnFilterAll, 
    $wpf_btnFilterActive, 
    $wpf_btnFilterLicGuests, 
    $wpf_btnFilterLicensed, 
    $wpf_btnFilterGuests, 
    $wpf_btnFilterUnlicensed, 
    $wpf_btnFilterDisabled, 
    $wpf_btnFilterInactive90d, 
    $wpf_btnFilterNever
)

function Set-PresetButtonActive($ActiveButton) {
    foreach ($btn in $Script:PresetButtons) {
        if ($null -ne $btn) {
            $btn.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushSecBtnBg')
            $btn.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushTextMuted')
            $btn.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
            $btn.BorderThickness = 1
        }
    }
    if ($null -ne $ActiveButton) {
        $ActiveButton.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccent')
        $ActiveButton.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushAccentFg')
        $ActiveButton.BorderThickness = 0
    }
}

$Script:SubLicButtons = @($wpf_btnSubLicAll, $wpf_btnSubLicYes, $wpf_btnSubLicNo)
$Script:SubTypeButtons = @($wpf_btnSubTypeAll, $wpf_btnSubTypeMembers, $wpf_btnSubTypeGuests, $wpf_btnSubTypeLicGuest)
$Script:SubAccButtons = @($wpf_btnSubAccAll, $wpf_btnSubAccEnabled, $wpf_btnSubAccDisabled)
$Script:SubActButtons = @($wpf_btnSubActAll, $wpf_btnSubActActive, $wpf_btnSubActInactive, $wpf_btnSubActNever)

function Set-SubButtonGroupActive($ButtonGroup, $ActiveBtn) {
    foreach ($btn in $ButtonGroup) {
        if ($null -ne $btn) {
            $btn.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushSecBtnBg')
            $btn.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushTextMuted')
            $btn.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushCardBorder')
            $btn.BorderThickness = 1
        }
    }
    if ($null -ne $ActiveBtn) {
        $ActiveBtn.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'BrushAccentTint')
        $ActiveBtn.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'BrushTextPrimary')
        $ActiveBtn.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'BrushAccent')
        $ActiveBtn.BorderThickness = 1
    }
}

function Reset-AllDirectoryFilters {
    $Script:ActivePreset = "all"
    $Script:FilterLicense = "all"
    $Script:FilterAccount = "all"
    $Script:FilterUserType = "all"
    $Script:FilterActivity = "all"
    $Script:SearchQuery = ""
    if ($wpf_txtSearch) { $wpf_txtSearch.Text = "" }
    
    Set-PresetButtonActive $wpf_btnFilterAll
    Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicAll
    Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeAll
    Set-SubButtonGroupActive $Script:SubAccButtons $wpf_btnSubAccAll
    Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActAll
    Filter-DataGrid
}

# Preset Button Handlers
if ($wpf_btnFilterAll) {
    $wpf_btnFilterAll.Add_Click({
        Reset-AllDirectoryFilters
    })
}

if ($wpf_btnFilterActive) {
    $wpf_btnFilterActive.Add_Click({
        $Script:ActivePreset = "active"
        Set-PresetButtonActive $wpf_btnFilterActive
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActActive
        Filter-DataGrid
    })
}

if ($wpf_btnFilterLicGuests) {
    $wpf_btnFilterLicGuests.Add_Click({
        $Script:ActivePreset = "licensed_guests"
        Set-PresetButtonActive $wpf_btnFilterLicGuests
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeLicGuest
        Filter-DataGrid
    })
}

if ($wpf_btnFilterLicensed) {
    $wpf_btnFilterLicensed.Add_Click({
        $Script:ActivePreset = "licensed"
        Set-PresetButtonActive $wpf_btnFilterLicensed
        Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicYes
        Filter-DataGrid
    })
}

if ($wpf_btnFilterGuests) {
    $wpf_btnFilterGuests.Add_Click({
        $Script:ActivePreset = "guests"
        Set-PresetButtonActive $wpf_btnFilterGuests
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeGuests
        Filter-DataGrid
    })
}

if ($wpf_btnFilterUnlicensed) {
    $wpf_btnFilterUnlicensed.Add_Click({
        $Script:ActivePreset = "unlicensed"
        Set-PresetButtonActive $wpf_btnFilterUnlicensed
        Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicNo
        Filter-DataGrid
    })
}

if ($wpf_btnFilterDisabled) {
    $wpf_btnFilterDisabled.Add_Click({
        $Script:ActivePreset = "disabled"
        Set-PresetButtonActive $wpf_btnFilterDisabled
        Set-SubButtonGroupActive $Script:SubAccButtons $wpf_btnSubAccDisabled
        Filter-DataGrid
    })
}

if ($wpf_btnFilterInactive90d) {
    $wpf_btnFilterInactive90d.Add_Click({
        $Script:ActivePreset = "inactive90d"
        Set-PresetButtonActive $wpf_btnFilterInactive90d
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActInactive
        Filter-DataGrid
    })
}

if ($wpf_btnFilterNever) {
    $wpf_btnFilterNever.Add_Click({
        $Script:ActivePreset = "never"
        Set-PresetButtonActive $wpf_btnFilterNever
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActNever
        Filter-DataGrid
    })
}

# Sub-Filter Additive Handlers
if ($wpf_btnSubLicAll) {
    $wpf_btnSubLicAll.Add_Click({
        $Script:FilterLicense = "all"
        Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicAll
        Filter-DataGrid
    })
}
if ($wpf_btnSubLicYes) {
    $wpf_btnSubLicYes.Add_Click({
        $Script:FilterLicense = "licensed"
        Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicYes
        Filter-DataGrid
    })
}
if ($wpf_btnSubLicNo) {
    $wpf_btnSubLicNo.Add_Click({
        $Script:FilterLicense = "unlicensed"
        Set-SubButtonGroupActive $Script:SubLicButtons $wpf_btnSubLicNo
        Filter-DataGrid
    })
}

if ($wpf_btnSubTypeAll) {
    $wpf_btnSubTypeAll.Add_Click({
        $Script:FilterUserType = "all"
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeAll
        Filter-DataGrid
    })
}
if ($wpf_btnSubTypeMembers) {
    $wpf_btnSubTypeMembers.Add_Click({
        $Script:FilterUserType = "members"
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeMembers
        Filter-DataGrid
    })
}
if ($wpf_btnSubTypeGuests) {
    $wpf_btnSubTypeGuests.Add_Click({
        $Script:FilterUserType = "guests"
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeGuests
        Filter-DataGrid
    })
}
if ($wpf_btnSubTypeLicGuest) {
    $wpf_btnSubTypeLicGuest.Add_Click({
        $Script:FilterUserType = "licensed_guests"
        Set-SubButtonGroupActive $Script:SubTypeButtons $wpf_btnSubTypeLicGuest
        Filter-DataGrid
    })
}

if ($wpf_btnSubAccAll) {
    $wpf_btnSubAccAll.Add_Click({
        $Script:FilterAccount = "all"
        Set-SubButtonGroupActive $Script:SubAccButtons $wpf_btnSubAccAll
        Filter-DataGrid
    })
}
if ($wpf_btnSubAccEnabled) {
    $wpf_btnSubAccEnabled.Add_Click({
        $Script:FilterAccount = "enabled"
        Set-SubButtonGroupActive $Script:SubAccButtons $wpf_btnSubAccEnabled
        Filter-DataGrid
    })
}
if ($wpf_btnSubAccDisabled) {
    $wpf_btnSubAccDisabled.Add_Click({
        $Script:FilterAccount = "disabled"
        Set-SubButtonGroupActive $Script:SubAccButtons $wpf_btnSubAccDisabled
        Filter-DataGrid
    })
}

if ($wpf_btnSubActAll) {
    $wpf_btnSubActAll.Add_Click({
        $Script:FilterActivity = "all"
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActAll
        Filter-DataGrid
    })
}
if ($wpf_btnSubActActive) {
    $wpf_btnSubActActive.Add_Click({
        $Script:FilterActivity = "active"
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActActive
        Filter-DataGrid
    })
}
if ($wpf_btnSubActInactive) {
    $wpf_btnSubActInactive.Add_Click({
        $Script:FilterActivity = "inactive90d"
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActInactive
        Filter-DataGrid
    })
}
if ($wpf_btnSubActNever) {
    $wpf_btnSubActNever.Add_Click({
        $Script:FilterActivity = "never"
        Set-SubButtonGroupActive $Script:SubActButtons $wpf_btnSubActNever
        Filter-DataGrid
    })
}

if ($wpf_btnResetFilters) {
    $wpf_btnResetFilters.Add_Click({
        Reset-AllDirectoryFilters
    })
}

# Exporting: Helper functions for Export Actions
function Invoke-CSVExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.csv"
        
        $ExcludedUPNs = @()
        $HasTestOrAdmin = $Script:GridRows | Where-Object { $_.VerificationText -in @("MSP Verification Account", "Test/Admin Account") }
        if ($HasTestOrAdmin) {
            $DlgResponse = Show-VerificationSelectorDialog -Users $HasTestOrAdmin
            if ($DlgResponse.Result -ne "Proceed") {
                Log-ToTerminal "CSV Export cancelled by user." "Warning"
                return
            }
            $ExcludedUPNs = $DlgResponse.ExcludedUPNs
        }
        
        $RowsToExport = $Script:GridRows | Where-Object { $_.UserPrincipalName -notin $ExcludedUPNs }
        $ExportRows = foreach ($row in $RowsToExport) {
            if ($Script:UserRole -eq "ServiceDesk") {
                [PSCustomObject]@{
                    "Display Name"        = $row.DisplayName
                    "User Principal Name" = $row.UserPrincipalName
                    "User Type"           = $row.UserType
                    "Assigned Licenses"   = $row.AssignedLicenses
                    "Last Sign-In Date"   = $row.LastSignInDate
                    "Status"              = $row.StatusText
                    "Account Status"      = $row.AccountStatusText
                    "Verification"        = $row.VerificationText
                }
            } else {
                [PSCustomObject]@{
                    "Display Name"              = $row.DisplayName
                    "User Principal Name"       = $row.UserPrincipalName
                    "User Type"                 = $row.UserType
                    "Assigned Licenses"         = $row.AssignedLicenses
                    "Last Sign-In Date"         = $row.LastSignInDate
                    "Status"                    = $row.StatusText
                    "Account Status"            = $row.AccountStatusText
                    "Wasted Cost (Accumulated)" = $row.WastedCostText
                    "Potential Monthly Savings" = $row.MonthlySavingsText
                    "Recommendation"            = $row.Recommendation
                    "Verification"              = $row.VerificationText
                }
            }
        }
        $ExportRows | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
        
        $MsgSuffix = if ($ExcludedUPNs.Count -gt 0) { " (Excluded $($ExcludedUPNs.Count) Test/Admin account(s))" } else { "" }
        Log-ToTerminal "CSV report auto-saved to: $Path$MsgSuffix" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("CSV successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "CSV Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("CSV Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-ExcelExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    $Module = Get-Module -ListAvailable -Name "ImportExcel"
    if (-not $Module) {
        $Msg = "The 'ImportExcel' module is required to generate Excel (.xlsx) files but was not found.`n`nWould you like the application to automatically download and install it from the PowerShell Gallery?"
        $Result = [System.Windows.MessageBox]::Show($Msg, "Module Missing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            Log-ToTerminal "ImportExcel module is missing. Initiating auto-installation..." "Warning"
            Update-UI
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Log-ToTerminal "Downloading and installing 'ImportExcel' from PSGallery..." "Info"
                Update-UI
                Install-Module -Name "ImportExcel" -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                Log-ToTerminal "ImportExcel installed successfully!" "Success"
                Update-UI
            } catch {
                Log-ToTerminal "Failed to install ImportExcel: $_" "Error"
                [System.Windows.MessageBox]::Show("Failed to install 'ImportExcel' automatically:`n$_`n`nPlease install it manually: Install-Module ImportExcel -Scope CurrentUser", "Installation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
        } else {
            return
        }
    }
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
    } catch {
        Log-ToTerminal "Failed to load 'ImportExcel' module: $_" "Error"
        [System.Windows.MessageBox]::Show("Failed to load 'ImportExcel' module:`n$_", "Load Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.xlsx"
        
        Log-ToTerminal "Exporting grid data to Excel..." "Info"
        Update-UI
        
        $ExcludedUPNs = @()
        $HasTestOrAdmin = $Script:GridRows | Where-Object { $_.VerificationText -in @("MSP Verification Account", "Test/Admin Account") }
        if ($HasTestOrAdmin) {
            $DlgResponse = Show-VerificationSelectorDialog -Users $HasTestOrAdmin
            if ($DlgResponse.Result -ne "Proceed") {
                Log-ToTerminal "Excel Export cancelled by user." "Warning"
                return
            }
            $ExcludedUPNs = $DlgResponse.ExcludedUPNs
        }
        
        $FilteredGridRows = @($Script:GridRows | Where-Object { $_.UserPrincipalName -notin $ExcludedUPNs })
        
        $ExportRows = @(foreach ($row in $FilteredGridRows) {
            if ($Script:UserRole -eq "ServiceDesk") {
                [PSCustomObject]@{
                    "Display Name"        = $row.DisplayName
                    "User Principal Name" = $row.UserPrincipalName
                    "User Type"           = $row.UserType
                    "Assigned Licenses"   = $row.AssignedLicenses
                    "Last Sign-In Date"   = $row.LastSignInDate
                    "Status"              = $row.StatusText
                    "Account Status"      = $row.AccountStatusText
                    "Verification"        = $row.VerificationText
                }
            } else {
                [PSCustomObject]@{
                    "Display Name"              = $row.DisplayName
                    "User Principal Name"       = $row.UserPrincipalName
                    "User Type"                 = $row.UserType
                    "Assigned Licenses"         = $row.AssignedLicenses
                    "Last Sign-In Date"         = $row.LastSignInDate
                    "Status"                    = $row.StatusText
                    "Account Status"            = $row.AccountStatusText
                    "Wasted Cost (Accumulated)" = $row.WastedCostText
                    "Potential Monthly Savings" = $row.MonthlySavingsText
                    "Recommendation"            = $row.Recommendation
                    "Verification"              = $row.VerificationText
                }
            }
        })
        
        if ($ExportRows.Count -eq 0) {
            [System.Windows.MessageBox]::Show("There are no rows to export.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Lock check: try opening file for writing if it already exists
        if (Test-Path $Path) {
            try {
                $Stream = [System.IO.File]::OpenWrite($Path)
                $Stream.Close()
            } catch {
                throw "The Excel file is currently open or locked by another process (e.g. Microsoft Excel). Please close it and try again."
            }
        }
        
        $Excel = $ExportRows | Export-Excel -Path $Path -PassThru -AutoSize -WorksheetName "M365 Licensed Users"
        $Sheet = $Excel.Workbook.Worksheets["M365 Licensed Users"]
        
        $ColsList = if ($Script:UserRole -eq "ServiceDesk") { @("A", "B", "C", "D", "E", "F", "G") } else { @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J") }
        
        for ($i = 0; $i -lt $FilteredGridRows.Count; $i++) {
            $RowIndex = $i + 2 # row 1 is header
            $Cat = $FilteredGridRows[$i].StatusCategory
            
            $BgColor = $null
            $FontColor = "White"
            
            switch ($Cat) {
                "Active (<=30d)"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#10b981") }
                "Inactive (30-90d)"   { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#eab308"); $FontColor = "Black" } # Black text on yellow for legibility
                "Inactive (90-365d)"  { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#3B82F6") }
                "Inactive (>1yr)"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#ef4444") }
                "Never Logged In"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#64748b") }
            }
            
            if ($BgColor) {
                foreach ($Col in $ColsList) {
                    $Cell = $Sheet.Cells["$Col$RowIndex"]
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($BgColor.A, $BgColor.R, $BgColor.G, $BgColor.B)
                    if ($FontColor) {
                        $DColor = if ($FontColor -eq "Black") { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
                        $Cell.Style.Font.Color.SetColor($DColor.A, $DColor.R, $DColor.G, $DColor.B)
                    }
                }
                
                # E is Status
                Set-ExcelRange -Worksheet $Sheet -Range "E$RowIndex" -Bold -HorizontalAlignment Center
                # F is Account Status
                Set-ExcelRange -Worksheet $Sheet -Range "F$RowIndex" -HorizontalAlignment Center
                
                if ($Script:UserRole -eq "ServiceDesk") {
                    # G is Verification
                    Set-ExcelRange -Worksheet $Sheet -Range "G$RowIndex" -HorizontalAlignment Center
                } else {
                    # G is Wasted Cost, H is Monthly Savings
                    Set-ExcelRange -Worksheet $Sheet -Range "G$RowIndex" -HorizontalAlignment Right
                    Set-ExcelRange -Worksheet $Sheet -Range "H$RowIndex" -HorizontalAlignment Right
                    # I is Recommendation
                    Set-ExcelRange -Worksheet $Sheet -Range "I$RowIndex" -HorizontalAlignment Left
                    # J is Verification
                    Set-ExcelRange -Worksheet $Sheet -Range "J$RowIndex" -HorizontalAlignment Center
                }
            }
        }
        
        # Add the second worksheet for unassigned breakdown if data is available
        if ($null -ne $Script:UnassignedLicenses -and $Script:UnassignedLicenses.Count -gt 0) {
            $UnassignedExport = foreach ($lic in $Script:UnassignedLicenses) {
                if ($Script:UserRole -eq "ServiceDesk") {
                    [PSCustomObject]@{
                        "Product License" = $lic.SkuPartName
                        "Sku ID"          = $lic.SkuId
                        "Active Units"    = $lic.ActiveUnits
                        "Consumed Units"  = $lic.ConsumedUnits
                        "Unassigned"      = $lic.UnassignedUnits
                    }
                } else {
                    [PSCustomObject]@{
                        "Product License" = $lic.SkuPartName
                        "Sku ID"          = $lic.SkuId
                        "Active Units"    = $lic.ActiveUnits
                        "Consumed Units"  = $lic.ConsumedUnits
                        "Unassigned"      = $lic.UnassignedUnits
                        "Monthly Cost"    = $lic.MonthlyPrice
                        "Wasted Cost/mo"  = $lic.WastedCost
                    }
                }
            }
            $Excel = $UnassignedExport | Export-Excel -ExcelPackage $Excel -PassThru -AutoSize -WorksheetName "Unassigned Licenses Breakdown"
            $UnusedSheet = $Excel.Workbook.Worksheets["Unassigned Licenses Breakdown"]
            
            # Style the header row in the new worksheet
            $HeaderBg = [System.Drawing.ColorTranslator]::FromHtml("#1c1c1e")
            $UnusedCols = if ($Script:UserRole -eq "ServiceDesk") { @("A", "B", "C", "D", "E") } else { @("A", "B", "C", "D", "E", "F", "G") }
            foreach ($Col in $UnusedCols) {
                $Cell = $UnusedSheet.Cells["$Col`1"]
                $Cell.Style.Font.Bold = $true
                $Cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                if ($HeaderBg) {
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($HeaderBg.A, $HeaderBg.R, $HeaderBg.G, $HeaderBg.B)
                }
                $Cell.Style.Font.Color.SetColor(255, 255, 255, 255) # White
            }
            
            # Alignments & formats
            for ($i = 0; $i -lt $UnassignedExport.Count; $i++) {
                $RowIndex = $i + 2
                
                # E is Unassigned
                $CellE = $UnusedSheet.Cells["E$RowIndex"]
                $CellE.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                if ($UnassignedExport[$i].Unassigned -gt 0) {
                    $CellE.Style.Font.Bold = $true
                    $CellE.Style.Font.Color.SetColor(255, 0, 128, 0) # Green (0, 128, 0)
                } elseif ($UnassignedExport[$i].Unassigned -lt 0) {
                    $CellE.Style.Font.Bold = $true
                    $CellE.Style.Font.Color.SetColor(255, 255, 0, 0) # Red (255, 0, 0)
                }
                
                if ($Script:UserRole -ne "ServiceDesk") {
                    Set-ExcelRange -Worksheet $UnusedSheet -Range "F$RowIndex" -NumberFormat "$#,##0.00" -HorizontalAlignment Right
                    
                    # G is Wasted Cost/mo
                    $CellG = $UnusedSheet.Cells["G$RowIndex"]
                    $CellG.Style.Numberformat.Format = "$#,##0.00"
                    $CellG.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
                    if ($UnassignedExport[$i]."Wasted Cost/mo" -gt 0) {
                        $CellG.Style.Font.Bold = $true
                        $CellG.Style.Font.Color.SetColor(255, 255, 0, 0) # Red
                    } elseif ($UnassignedExport[$i]."Wasted Cost/mo" -lt 0) {
                        $CellG.Style.Font.Bold = $true
                        $CellG.Style.Font.Color.SetColor(255, 0, 128, 0) # Green
                    }
                }
            }
        }
        
        Close-ExcelPackage -ExcelPackage $Excel
        
        $MsgSuffix = if ($ExcludedUPNs.Count -gt 0) { " (Excluded $($ExcludedUPNs.Count) Test/Admin account(s))" } else { "" }
        Log-ToTerminal "Excel report auto-saved to: $Path$MsgSuffix" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("Excel successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "Excel Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Excel Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Export-WpfPng($Element, $FilePath) {
    # WPF layout bounds
    $width = $Element.ActualWidth
    $height = $Element.ActualHeight
    
    if ($width -le 0 -or $height -le 0) {
        $width = 900
        $height = 650
    }
    
    # Render element to bitmap context
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
        [int]$width, [int]$height, 96, 96, 
        [System.Windows.Media.PixelFormats]::Pbgra32
    )
    $rtb.Render($Element)
    
    # Save bitmap as PNG image
    $pngEncoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $pngEncoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    
    $fs = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Create)
    $pngEncoder.Save($fs)
    $fs.Close()
}

function Invoke-PNGExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.png"
        
        Export-WpfPng $wpf_pageDashboard $Path
        Log-ToTerminal "PNG snapshot auto-saved to: $Path" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("PNG successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "PNG Snapshot failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Snapshot Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-PDFExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_ExecutiveSummary_${CleanTenant}_${DateStr}.html"
        
        Log-ToTerminal "Generating print-ready Executive Summary HTML..." "Info"
        Update-UI
        
        # Gather metrics
        $Total = $Script:GridRows.Count
        $Active = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Active (<=30d)" }).Count
        $Inactive = ($Script:GridRows | Where-Object { $_.DaysSince -gt 180 }).Count
        
        $TotalSavings = 0.00
        $TotalMonthlySavings = 0.00
        foreach ($row in $Script:GridRows) {
            $TotalSavings += $row.WastedCost
            $TotalMonthlySavings += $row.MonthlySavings
        }
        
        # Recommendations list (high priority action items)
        $RecRows = $Script:GridRows | Where-Object { $_.Recommendation -ne "-" }
        $RecTableHtml = ""
        if ($RecRows.Count -gt 0) {
            foreach ($row in $RecRows) {
                $RecTableHtml += @"
            <tr>
                <td style="font-weight: 500;">$($row.DisplayName)</td>
                <td style="font-family: monospace;">$($row.UserPrincipalName)</td>
                <td>$($row.LastSignInDate)</td>
                <td style="color: #ef4444; font-weight: 600;">$($row.WastedCostText)</td>
                <td style="color: #10b981; font-weight: 600;">$($row.MonthlySavingsText)</td>
                <td style="color: #3B82F6; font-weight: 600;">$($row.Recommendation)</td>
            </tr>
"@
            }
        } else {
            $RecTableHtml = "<tr><td colspan='6' style='text-align: center; color: #9ca3af;'>No pending recommendations. All users active!</td></tr>"
        }
        
        # Full users list table
        $UserTableHtml = ""
        foreach ($row in $Script:GridRows) {
            $StatusStyle = switch ($row.StatusCategory) {
                "Active (<=30d)"     { "background-color: #d1fae5; color: #065f46;" }
                "Inactive (30-90d)"   { "background-color: #fef9c3; color: #854d0e;" }
                "Inactive (90-365d)"  { "background-color: #ffedd5; color: #9a3412;" }
                "Inactive (>1yr)"     { "background-color: #fee2e2; color: #991b1b;" }
                "Never Logged In"     { "background-color: #f3f4f6; color: #374151;" }
                default              { "background-color: #f3f4f6; color: #374151;" }
            }
            $UserTableHtml += @"
            <tr>
                <td>$($row.DisplayName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($row.UserPrincipalName)</td>
                <td style="text-align: center; font-size: 11px; font-weight: 500;">$($row.UserType)</td>
                <td style="font-size: 11px;">$($row.AssignedLicenses)</td>
                <td>$($row.LastSignInDate)</td>
                <td style="text-align: center;"><span style="padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; $StatusStyle">$($row.StatusText)</span></td>
                <td style="text-align: center; font-size: 11px;">$($row.AccountStatusText)</td>
                <td style="text-align: right; font-weight: 500;">$($row.WastedCostText)</td>
                <td style="text-align: right; font-weight: 500;">$($row.MonthlySavingsText)</td>
                <td style="font-size: 11px; color: #3B82F6;">$($row.Recommendation)</td>
            </tr>
"@
        }
        
        # Unassigned licenses inventory table
        $UnassignedTableHtml = ""
        if ($null -ne $Script:UnassignedLicenses -and $Script:UnassignedLicenses.Count -gt 0) {
            $UnassignedRowsHtml = ""
            foreach ($lic in $Script:UnassignedLicenses) {
                # Format counts and colors
                $UnassignedStyle = if ($lic.UnassignedUnits -gt 0) { "color: #10b981; font-weight: 600;" } elseif ($lic.UnassignedUnits -lt 0) { "color: #ef4444; font-weight: 600;" } else { "color: #9ca3af;" }
                
                if ($Script:UserRole -eq "ServiceDesk") {
                    $UnassignedRowsHtml += @"
            <tr>
                <td style="font-weight: 500;">$($lic.SkuPartName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($lic.SkuId)</td>
                <td style="text-align: center;">$($lic.ActiveUnits)</td>
                <td style="text-align: center;">$($lic.ConsumedUnits)</td>
                <td style="text-align: center; $UnassignedStyle">$($lic.UnassignedUnits)</td>
            </tr>
"@
                } else {
                    $WasteStyle = if ($lic.WastedCost -gt 0) { "color: #ef4444; font-weight: 600;" } elseif ($lic.WastedCost -lt 0) { "color: #10b981; font-weight: 600;" } else { "color: #9ca3af;" }
                    $UnassignedRowsHtml += @"
            <tr>
                <td style="font-weight: 500;">$($lic.SkuPartName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($lic.SkuId)</td>
                <td style="text-align: center;">$($lic.ActiveUnits)</td>
                <td style="text-align: center;">$($lic.ConsumedUnits)</td>
                <td style="text-align: center; $UnassignedStyle">$($lic.UnassignedUnits)</td>
                <td style="text-align: right;">$($lic.MonthlyPriceText)</td>
                <td style="text-align: right; $WasteStyle">$($lic.WastedCostText)</td>
            </tr>
"@
                }
            }
            
            $TableHeaderHtml = if ($Script:UserRole -eq "ServiceDesk") {
                @"
            <tr>
                <th>Product License</th>
                <th>Sku ID</th>
                <th style="text-align: center;">Active Units</th>
                <th style="text-align: center;">Consumed Units</th>
                <th style="text-align: center;">Unassigned Units</th>
            </tr>
"@
            } else {
                @"
            <tr>
                <th>Product License</th>
                <th>Sku ID</th>
                <th style="text-align: center;">Active Units</th>
                <th style="text-align: center;">Consumed Units</th>
                <th style="text-align: center;">Unassigned Units</th>
                <th style="text-align: right;">Monthly Cost</th>
                <th style="text-align: right;">Wasted Cost/mo</th>
            </tr>
"@
            }
            
            $UnassignedTableHtml = @"
    <div class="section-title">Unassigned Licenses (Pool Waste) Breakdown</div>
    <table>
        <thead>
            $TableHeaderHtml
        </thead>
        <tbody>
            $UnassignedRowsHtml
        </tbody>
    </table>
"@
        }
        
        # Build HTML content
        $HtmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>M365 User License Audit - Executive Summary</title>
    <style>
        body {
            font-family: 'Segoe UI', -apple-system, sans-serif;
            color: #1f2937;
            background-color: #ffffff;
            margin: 0;
            padding: 30px;
            line-height: 1.5;
        }
        .header {
            border-bottom: 3px solid #3B82F6;
            padding-bottom: 20px;
            margin-bottom: 30px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }
        .header h1 {
            margin: 0;
            font-size: 24px;
            color: #111827;
        }
        .header p {
            margin: 5px 0 0 0;
            color: #4b5563;
            font-size: 14px;
        }
        .meta-box {
            text-align: right;
            font-size: 12px;
            color: #6b7280;
        }
        .meta-box strong {
            color: #111827;
        }
        .kpi-container {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }
        .kpi-card {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 15px;
            background-color: #f9fafb;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .kpi-card .title {
            font-size: 10px;
            font-weight: 700;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .kpi-card .value {
            font-size: 20px;
            font-weight: 700;
            margin-top: 5px;
            color: #111827;
        }
        .kpi-card .sub {
            font-size: 11px;
            color: #9ca3af;
            margin-top: 4px;
        }
        .section-title {
            font-size: 16px;
            font-weight: 700;
            color: #111827;
            margin-top: 30px;
            margin-bottom: 15px;
            border-bottom: 1px solid #e5e7eb;
            padding-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin-bottom: 25px;
        }
        th {
            background-color: #f3f4f6;
            color: #374151;
            text-align: left;
            font-weight: 600;
            padding: 8px 12px;
            border-bottom: 2px solid #e5e7eb;
        }
        td {
            padding: 8px 12px;
            border-bottom: 1px solid #f3f4f6;
            color: #4b5563;
        }
        tr:nth-child(even) td {
            background-color: #fafafa;
        }
        .footer {
            margin-top: 50px;
            border-top: 1px solid #e5e7eb;
            padding-top: 15px;
            text-align: center;
            font-size: 10px;
            color: #9ca3af;
        }
        @media print {
            body { padding: 0; }
            .kpi-card { box-shadow: none; }
            button { display: none; }
        }
    </style>
</head>
<body>

    <div class="header">
        <div>
            <h1>M365 User License Audit</h1>
            <p>Executive License Optimization & Cost Summary</p>
        </div>
        <div class="meta-box">
            Tenant Domain: <strong>$((Get-TenantName))</strong><br>
            Generated Date: <strong>$((Get-Date).ToString("dd/MM/yyyy HH:mm"))</strong><br>
            Session Mode: <strong>$($Script:SessionMode)</strong>
        </div>
    </div>

    <div class="kpi-container">
        <div class="kpi-card" style="border-left: 3px solid #3B82F6;">
            <div class="title">Total Licensed</div>
            <div class="value">$Total</div>
            <div class="sub">Active Entra accounts</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #10b981;">
            <div class="title">Active (&lt;=30d)</div>
            <div class="value">$Active</div>
            <div class="sub">Logging in regularly</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Inactive (&gt;180d)</div>
            <div class="value">$Inactive</div>
            <div class="sub">Wasting license fees</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Wasted Money</div>
            <div class="value" style="color: #ef4444;">£$("{0:N2}" -f $TotalSavings)</div>
            <div class="sub">Total historical loss</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #10b981;">
            <div class="title">Monthly Savings</div>
            <div class="value" style="color: #10b981;">£$("{0:N2}" -f $TotalMonthlySavings)</div>
            <div class="sub">Future monthly reclaim</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Pool Waste</div>
            <div class="value" style="color: #ef4444;">£$("{0:N2}" -f $Script:TotalPoolWaste)</div>
            <div class="sub">$Script:TotalUnassignedCount unassigned units</div>
        </div>
    </div>

    <div class="section-title">High-Priority Optimization Actions</div>
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>User Principal Name</th>
                <th>Last Sign-In Date</th>
                <th>Wasted Cost</th>
                <th>Monthly Savings</th>
                <th>Remediation Recommendation</th>
            </tr>
        </thead>
        <tbody>
            $RecTableHtml
        </tbody>
    </table>

    $UnassignedTableHtml

    <div class="section-title">Full User License Inventory</div>
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>User Principal Name</th>
                <th style="text-align: center;">Type</th>
                <th>Assigned Licenses</th>
                <th>Last Sign-In Date</th>
                <th>Status</th>
                <th style="text-align: center;">Account Status</th>
                <th style="text-align: right;">Wasted Cost</th>
                <th style="text-align: right;">Monthly Savings</th>
                <th>Recommendation</th>
            </tr>
        </thead>
        <tbody>
            $UserTableHtml
        </tbody>
    </table>

    <div class="footer">
        Generated by M365 PowerShell Admin Console &copy; 2026 Contoso. All rights reserved.
    </div>

    <script>
        window.onload = function() {
            setTimeout(function() {
                window.print();
            }, 500);
        }
    </script>
</body>
</html>
"@
        
        $HtmlTemplate | Out-File -FilePath $Path -Force -Encoding utf8
        Log-ToTerminal "HTML report successfully generated at: $Path" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("Executive HTML summary report exported successfully to:`n$Path`n`nWould you like to open it in your browser to print/save as PDF?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            Start-Process $Path
        }
    } catch {
        Log-ToTerminal "PDF/HTML Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("PDF/HTML Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-EmailDraftGeneration {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No user data available to generate an email draft. Please connect and run the audit first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Helper to detect if a specific license is a free / non-commercial tier
    $IsFreeLicense = {
        param([string]$licName)
        if ([string]::IsNullOrWhiteSpace($licName) -or $licName -eq "None (Unlicensed)") { return $true }
        $n = $licName.Trim()
        if ($n -match "Free" -or $n -match "Exploratory" -or $n -match "Trial" -or $n -match "Developer" -or $n -match "Dev" -or $n -match "Store for Business" -or $n -match "Service Business" -or $n -match "Fabric" -or $n -match "Eval") {
            return $true
        }
        $price = Get-LicenseMonthlyPrice $n
        return ($price -le 0.00)
    }

    # Helper to detect accounts holding paid commercial subscriptions (strictly excluding ALL free/trial/exploratory tiers)
    $IsPaidLicenseRow = {
        param($row)
        if (-not $row.AssignedLicenses -or $row.AssignedLicenses -eq "None (Unlicensed)") { return $false }
        $paidPlans = $row.AssignedLicenses.Split(',') | ForEach-Object { $_.Trim() } | Where-Object {
            $_ -and (-not (& $IsFreeLicense $_))
        }
        return ($paidPlans.Count -gt 0)
    }

    # Calculate counts strictly based on paid commercial licensing
    $TotalAudited = $Script:GridRows.Count
    $PaidLicensedRows = $Script:GridRows | Where-Object { & $IsPaidLicenseRow $_ }
    $PaidLicensedCount = $PaidLicensedRows.Count
    $GuestCount = ($Script:GridRows | Where-Object { $_.IsGuest }).Count
    $DisabledCount = ($Script:GridRows | Where-Object { -not $_.AccountEnabled }).Count
    $DisabledWithPaidLic = ($PaidLicensedRows | Where-Object { -not $_.AccountEnabled }).Count

    # Inactivity metrics evaluated on paid licensed accounts for genuine licensing optimization
    $Active = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Active (<=30d)" }).Count
    $Warning = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Inactive (30-90d)" }).Count
    $Inactive = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Inactive (90-365d)" }).Count
    $Critical = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Inactive (>1yr)" }).Count
    $Never = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Never Logged In" }).Count
    $NoP1P2 = ($PaidLicensedRows | Where-Object { $_.StatusCategory -eq "Requires Entra ID P1/P2" }).Count

    # Count accounts that hold ONLY free/complimentary licenses
    $FreeOnlyAccountsCount = ($Script:GridRows | Where-Object { 
        $_.AssignedLicenses -and 
        $_.AssignedLicenses -ne "None (Unlicensed)" -and 
        (-not (& $IsPaidLicenseRow $_))
    }).Count
    
    # Tenant details
    $TenantName = (Get-TenantName)
    if ([string]::IsNullOrWhiteSpace($TenantName) -or $TenantName -eq "Not Connected") {
        $TenantName = "[Tenant Name]"
    }
    
    # Dynamic savings call-to-action & tone based on user role
    $IsServiceDesk = ($Script:UserRole -eq "ServiceDesk")
    $CTA = ""
    $SubjectLine = ""
    $IntroText = ""
    $CaveatIntro = ""
    
    if ($IsServiceDesk) {
        $SubjectLine = "Subject: Microsoft 365 User License & Inactivity Audit - $TenantName"
        $IntroText = "This report provides a breakdown of your directory user accounts (licensed members, unlicensed accounts, and external guests) and their associated Microsoft 365 subscriptions. The spreadsheet is color-coded by login activity to help identify inactive or dormant user accounts:"
        $CaveatIntro = "Please note that this audit acts as a technical guide to highlight inactive user accounts across your directory. Before disabling or unassigning any licenses, we recommend cross-referencing this list against the following scenarios:"
        $CTA = "We recommend reviewing this list with your department heads to confirm which licenses can be safely unassigned or reallocated."
    } else {
        # Calculate monthly savings for Sales / Business role
        $TotalMonthlySavings = 0.00
        foreach ($row in $Script:GridRows) {
            $TotalMonthlySavings += $row.MonthlySavings
        }
        $SubjectLine = "Subject: Microsoft 365 License Audit & Optimization Report - $TenantName"
        $IntroText = "This report provides a breakdown of your directory user accounts (licensed members, unlicensed accounts, and external guests) and their associated Microsoft 365 subscriptions. The spreadsheet is color-coded by login activity to help identify potential licensing optimization opportunities:"
        $CaveatIntro = "Please note that this audit acts as a rough guide to highlight potential waste. Before disabling or unassigning any licenses, we recommend cross-referencing this list against the following scenarios:"
        $CTA = "We recommend reviewing this list with your department heads to confirm which licenses can be safely reclaimed to reduce your monthly subscription costs (potential savings of £{0:N2}/mo)." -f $TotalMonthlySavings
    }
    
    $IconPurple = [char]::ConvertFromUtf32(0x1F7E6)
    $IconGreen  = [char]::ConvertFromUtf32(0x1F7E9)
    $IconYellow = [char]::ConvertFromUtf32(0x1F7E8)
    $IconOrange = [char]::ConvertFromUtf32(0x1F7E7)
    $IconRed    = [char]::ConvertFromUtf32(0x1F7E5)
    $IconGray   = [char]::ConvertFromUtf32(0x2B1B)
    
    $StatusSummaryLines = [System.Collections.Generic.List[string]]::new()
    $StatusSummaryLines.Add("* Tenant Directory Summary: $TotalAudited total users audited ($PaidLicensedCount paid licensed accounts, $GuestCount guest accounts, $DisabledCount disabled accounts).")
    
    # Active paid subscriptions breakdown (strictly excluding all free licenses)
    $ProductCounts = [System.Collections.Generic.Dictionary[string, int]]::new()
    foreach ($row in $PaidLicensedRows) {
        $row.AssignedLicenses.Split(',') | ForEach-Object {
            $p = $_.Trim()
            if ($p -and (-not (& $IsFreeLicense $p))) {
                if ($ProductCounts.ContainsKey($p)) {
                    $ProductCounts[$p]++
                } else {
                    $ProductCounts[$p] = 1
                }
            }
        }
    }
    if ($ProductCounts.Count -gt 0) {
        $StatusSummaryLines.Add("* Active Paid Subscriptions Breakdown:")
        foreach ($k in ($ProductCounts.Keys | Sort-Object)) {
            $StatusSummaryLines.Add("  - ${k}: $($ProductCounts[$k]) assigned")
        }
    }

    if ($FreeOnlyAccountsCount -gt 0) {
        if ($IsServiceDesk) {
            $StatusSummaryLines.Add("* Complimentary Tiers: $FreeOnlyAccountsCount account(s) hold only free/exploratory licenses and are excluded from the paid license counts.")
        } else {
            $StatusSummaryLines.Add("* Complimentary Tiers: $FreeOnlyAccountsCount account(s) hold only free/exploratory licenses (£0.00/mo) and are excluded from the paid license counts.")
        }
    }

    if ($DisabledWithPaidLic -gt 0) {
        $StatusSummaryLines.Add("* [ACTION REQUIRED] $DisabledWithPaidLic disabled/blocked account(s) currently retain active paid subscriptions.")
    }
    if ($NoP1P2 -gt 0) {
        $StatusSummaryLines.Add("* $IconPurple Purple (Requires Entra ID P1/P2): $NoP1P2 user(s). Account sign-in timestamps require Entra ID P1/P2 licensing.")
    }
    $StatusSummaryLines.Add("* $IconGreen Green (Active <= 30 days): $Active user(s). Logging in regularly. No action required.")
    $StatusSummaryLines.Add("* $IconYellow Yellow (Warning 30-90 days): $Warning user(s). Moderate inactivity.")
    $StatusSummaryLines.Add("* $IconOrange Orange (Inactive 90-365 days): $Inactive user(s). High inactivity. Recommended for review.")
    $CandidateText = if ($IsServiceDesk) { "Strong candidates for license unassignment." } else { "Strong candidates for license removal." }
    $StatusSummaryLines.Add("* $IconRed Red (Critical > 1 year): $Critical user(s). Extremely high inactivity. $CandidateText")
    $StatusSummaryLines.Add("* $IconGray Gray (Never Logged In): $Never user(s). Accounts that have never registered an interactive login.")
    
    $StatusSummaryText = $StatusSummaryLines -join "`r`n"
    
    # Formulate email draft
    $EmailText = @"
$SubjectLine

Dear [Director Name],

Please find attached the latest Microsoft 365 User License Audit report for your tenant. 

$IntroText

$StatusSummaryText

Important Context & Caveats
$CaveatIntro
1.  Delegated Access: Some accounts (e.g., shared mailboxes or role accounts) are accessed purely via delegation by other users, which means the mailbox is active but the underlying account registers no direct logins. In some cases, but not all, this means the license can be freed up.
2.  Extended Leave: Users currently on maternity, paternity, or long-term sick leave will show as inactive, but their accounts and licenses must remain active during their absence.
3.  New Starters: Accounts set up for new staff members who never actually joined the business.
4.  Project Accounts & Integrations: Temporary accounts created for specific projects, testing, or historical third-party software integrations that are no longer in use.
5.  Untracked Leavers: There may be former staff members who have left the business, but we were not notified to offboard them or remove their licenses.
6.  Temporary Re-licensing: Accounts that were temporarily re-licensed to allow team members to retrieve files, but we were not notified to remove the license once the request was complete.

$CTA

Kind regards,

[Your Company Name] Service Desk
[Contact Information]
"@

    Show-EmailTemplateDialog -EmailText $EmailText
}

# Bind click events for directory grid page exporter buttons
$wpf_btnExportCSV.Add_Click({ Invoke-CSVExport })
$wpf_btnExportExcel.Add_Click({ Invoke-ExcelExport })
$wpf_btnExportPNG.Add_Click({ Invoke-PNGExport })
$wpf_btnExportPDF.Add_Click({ Invoke-PDFExport })
$wpf_btnGenEmail.Add_Click({ Invoke-EmailDraftGeneration })

# Bind click events for dashboard page exporter buttons to delegate to the same functions
$wpf_btnExportCSVDash.Add_Click({ Invoke-CSVExport })
$wpf_btnExportExcelDash.Add_Click({ Invoke-ExcelExport })
$wpf_btnExportPNGDash.Add_Click({ Invoke-PNGExport })
$wpf_btnExportPDFDash.Add_Click({ Invoke-PDFExport })
$wpf_btnGenEmailDash.Add_Click({ Invoke-EmailDraftGeneration })

# Clear Graph cache at startup to ensure a clean slate
Clear-MgGraphCache

# Initialize tab pages
Show-Page "Connection"

# 12. Show Window Frame Dialog
$Window.ShowDialog() | Out-Null
