#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Online Mobile Device Quarantine Release GUI Tool
.DESCRIPTION
    A desktop GUI tool to connect to Exchange Online, search and select licensed mailbox users,
    check their mobile device quarantine status, and release devices to the Allowed list.
.NOTES
    Script Name   : DeviceQuarantineReleaseApp.ps1
    Author        : JP
    Version       : 1.1.0
    Log Path      : C:\Logs\DeviceQuarantineRelease\DeviceQuarantineRelease.log
#>

# Enforce STA Threading for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
    exit
}

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

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
            BrushOkTint        = '#1F34D399'
            BrushWarn          = '#FBBF24'
            BrushWarnTint      = '#26FBBF24'
            BrushWarnBorder    = '#66FBBF24'
            BrushErr           = '#F87171'
            BrushGridHeaderBg  = '#0B1528'
            BrushGridRowBg     = '#091122'
            BrushGridAltRowBg  = '#0C172B'
            BrushGridSelectedBg = '#1E3A5F'
            BrushGridLine      = '#1A2B44'
            BrushDangerBg      = '#26EF4444'
            BrushDangerHover   = '#3CEF4444'
            BrushDangerBorder  = '#66EF4444'
            BrushDangerFg      = '#FCA5A5'
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
            BrushOkTint        = '#D1FAE5'
            BrushWarn          = '#B45309'
            BrushWarnTint      = '#FEF3C7'
            BrushWarnBorder    = '#F59E0B'
            BrushErr           = '#DC2626'
            BrushGridHeaderBg  = '#F1F5F9'
            BrushGridRowBg     = '#FFFFFF'
            BrushGridAltRowBg  = '#F8FAFC'
            BrushGridSelectedBg = '#DBEAFE'
            BrushGridLine      = '#E2E8F0'
            BrushDangerBg      = '#FEE2E2'
            BrushDangerHover   = '#FECACA'
            BrushDangerBorder  = '#FCA5A5'
            BrushDangerFg      = '#B91C1C'
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

# Resolve log directory (UK date format dd/MM/yy HH:mm:ss)
$script:LogDirectory = 'C:\Logs\DeviceQuarantineRelease'
try {
    if (-not (Test-Path -Path $script:LogDirectory)) {
        New-Item -ItemType Directory -Path $script:LogDirectory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $script:LogDirectory = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Logs\DeviceQuarantineRelease'
    if (-not (Test-Path -Path $script:LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $script:LogDirectory -Force -ErrorAction Stop | Out-Null
        } catch {}
    }
}
$script:LogFilePath = Join-Path -Path $script:LogDirectory -ChildPath ("DeviceQuarantineRelease_{0}.log" -f (Get-Date -Format 'dd-MM-yy'))

# Define XAML layout (Dynamic System Dark/Light Theme)
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Online Quarantine Release Manager v1.1"
        Height="720" Width="1060" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual" FontSize="13" Background="#091122" Foreground="{DynamicResource BrushTextPrimary}">
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
        <SolidColorBrush x:Key="BrushWarnTint" Color="#26FBBF24"/>
        <SolidColorBrush x:Key="BrushWarnBorder" Color="#66FBBF24"/>
        <SolidColorBrush x:Key="BrushErr" Color="#F87171"/>
        <SolidColorBrush x:Key="BrushGridHeaderBg" Color="#0B1528"/>
        <SolidColorBrush x:Key="BrushGridRowBg" Color="#091122"/>
        <SolidColorBrush x:Key="BrushGridAltRowBg" Color="#0C172B"/>
        <SolidColorBrush x:Key="BrushGridSelectedBg" Color="#1E3A5F"/>
        <SolidColorBrush x:Key="BrushGridLine" Color="#1A2B44"/>
        <SolidColorBrush x:Key="BrushDangerBg" Color="#26EF4444"/>
        <SolidColorBrush x:Key="BrushDangerHover" Color="#3CEF4444"/>
        <SolidColorBrush x:Key="BrushDangerBorder" Color="#66EF4444"/>
        <SolidColorBrush x:Key="BrushDangerFg" Color="#FCA5A5"/>

        <Style x:Key="HeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="19"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
        </Style>
        <Style x:Key="SubHeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
        </Style>

        <!-- Primary CTA Button -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
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

        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
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

        <!-- Action CTA Button -->
        <Style x:Key="ActionButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>

        <!-- Input TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="SelectionBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" CornerRadius="6" BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}" Background="{TemplateBinding Background}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ScrollBar -->
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

        <!-- Directory ListBox Style -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{DynamicResource BrushGridRowBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ListBoxItem">
                        <Setter Property="Padding" Value="10,9"/>
                        <Setter Property="Margin" Value="0,1"/>
                        <Setter Property="Background" Value="Transparent"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="BorderBrush" Value="Transparent"/>
                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ListBoxItem">
                                    <Border x:Name="Border" CornerRadius="5" Background="{TemplateBinding Background}"
                                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                            Padding="{TemplateBinding Padding}">
                                        <ContentPresenter/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="{DynamicResource BrushGridSelectedBg}"/>
                                            <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                                            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="130"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER BAR -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Exchange Online Quarantine Release Manager v1.1" Style="{StaticResource HeaderStyle}"/>
                    <TextBlock Text="Search Exchange Online mailboxes and release quarantined ActiveSync mobile devices." Style="{StaticResource SubHeaderStyle}" Margin="0,3,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnThemeToggle" Style="{StaticResource SecondaryButton}" Height="36" Margin="0,0,10,0" Padding="12,6" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock x:Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <Border x:Name="StatusBadge" Background="{DynamicResource BrushSecBtnBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="11,6" Margin="0,0,12,0" VerticalAlignment="Center">
                        <TextBlock x:Name="StatusText" Text="Disconnected" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center" FontWeight="SemiBold" FontSize="12"/>
                    </Border>
                    <Button x:Name="BtnConnect" Content="Connect to Exchange Online" Height="36" MinWidth="200"/>
                    <Button x:Name="BtnDisconnect" Content="Disconnect" Height="36" MinWidth="105" Margin="10,0,0,0" Style="{StaticResource SecondaryButton}" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN WORKSPACE -->
        <Grid Grid.Row="1" Margin="16,14,16,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="340" MinWidth="260"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT COLUMN: USER DIRECTORY -->
            <Border Grid.Column="0" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="14">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Local Filtering Search Box -->
                    <Grid Grid.Row="0" Margin="0,0,0,10">
                        <TextBox x:Name="TxtSearch" Height="36" VerticalContentAlignment="Center" IsEnabled="False"/>
                        <TextBlock IsHitTestVisible="False" Text="Filter mailbox users..." VerticalAlignment="Center" HorizontalAlignment="Left" Margin="12,0,0,0" Foreground="{DynamicResource BrushTextMuted}">
                            <TextBlock.Style>
                                <Style TargetType="{x:Type TextBlock}">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding Text, ElementName=TxtSearch}" Value="">
                                            <Setter Property="Visibility" Value="Visible"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </TextBlock.Style>
                        </TextBlock>
                    </Grid>

                    <!-- Mailbox User List -->
                    <ListBox Grid.Row="1" x:Name="LstUsers" Margin="0,0,0,10" IsEnabled="False">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <StackPanel>
                                    <TextBlock Text="{Binding DisplayName}" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13"/>
                                    <TextBlock Text="{Binding PrimarySmtpAddress}" Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5" Margin="0,2,0,0"/>
                                </StackPanel>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <!-- Fetch Directory Action -->
                    <Button Grid.Row="2" x:Name="BtnLoadUsers" Content="Load User Directory" Style="{StaticResource SecondaryButton}" Height="36" IsEnabled="False"/>
                </Grid>
            </Border>

            <!-- RIGHT COLUMN: SELECTION & DEVICE PANEL -->
            <Border Grid.Column="2" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="16">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Selected Target Header -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,14">
                        <TextBlock x:Name="TxtSelectedUserName" Text="No Mailbox Selected" FontSize="17" FontWeight="Bold" Foreground="{DynamicResource BrushTextPrimary}"/>
                        <TextBlock x:Name="TxtSelectedUserEmail" Text="Select a mailbox from the left panel to display quarantine status." FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}" Margin="0,3,0,0"/>
                    </StackPanel>

                    <!-- Quarantined Devices Container -->
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="Detected Quarantined Devices" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,0,0,8"/>

                        <ListBox Grid.Row="1" x:Name="LstDevices">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <Border BorderBrush="{DynamicResource BrushGridLine}" BorderThickness="0,0,0,1" Padding="4,6">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>

                                            <StackPanel>
                                                <TextBlock Text="{Binding DeviceModel}" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextPrimary}" FontSize="13.5"/>
                                                <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                                                    <TextBlock Text="Type: " Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5"/>
                                                    <TextBlock Text="{Binding DeviceType}" Foreground="{DynamicResource BrushTextSecondary}" FontSize="11.5" Margin="0,0,14,0"/>
                                                    <TextBlock Text="ID: " Foreground="{DynamicResource BrushTextMuted}" FontSize="11.5"/>
                                                    <TextBlock Text="{Binding DeviceId}" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="11.5"/>
                                                </StackPanel>
                                            </StackPanel>

                                            <Border Grid.Column="1" Background="{DynamicResource BrushWarnTint}" BorderBrush="{DynamicResource BrushWarnBorder}" BorderThickness="1" CornerRadius="4" Padding="8,4" VerticalAlignment="Center" Margin="12,0,0,0">
                                                <TextBlock Text="{Binding DeviceAccessState}" Foreground="{DynamicResource BrushWarn}" FontSize="11" FontWeight="Bold"/>
                                            </Border>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>

                        <!-- Empty/Feedback States overlaying the ListBox -->
                        <TextBlock Grid.Row="1" x:Name="TxtNoDevices" Text="No quarantined devices found for this user." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource BrushTextMuted}" Visibility="Collapsed"/>
                        <TextBlock Grid.Row="1" x:Name="TxtLoadingDevices" Text="Querying Mobile Devices..." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource BrushAccent}" FontWeight="SemiBold" Visibility="Collapsed"/>
                    </Grid>
                </Grid>
            </Border>
        </Grid>

        <!-- LIVE ACTIVITY CONSOLE -->
        <Border Grid.Row="2" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Margin="16,0,16,12" Padding="10,8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Activity Console" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                <TextBox Grid.Row="1" x:Name="TxtConsole" Background="Transparent" BorderThickness="0" Padding="0"
                         FontFamily="Cascadia Mono, Consolas" FontSize="11.5" Foreground="{DynamicResource BrushTextSecondary}"
                         IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" TextWrapping="Wrap"/>
            </Grid>
        </Border>

        <!-- SINGLE-ROW FOOTER -->
        <Border Grid.Row="3" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" Text="Log: C:\Logs\DeviceQuarantineRelease" Foreground="{DynamicResource BrushTextMuted}"
                           FontFamily="Cascadia Mono, Consolas" FontSize="11.5" VerticalAlignment="Center"
                           TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnCheckDevices" Content="Re-Scan Status" Style="{StaticResource SecondaryButton}" Height="36" MinWidth="140" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="BtnRelease" Content="Release Quarantined Devices" Style="{StaticResource ActionButton}" Height="36" MinWidth="220" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Read XAML layout
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
try {
    $Form = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "XAML parse error: $_" -ForegroundColor Red
    exit
}

Initialize-StandardGuiWindow -Window $Form

# Bind elements with Name/x:Name attributes to local PowerShell variables
$XAML.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
    $attr = $_.Attributes | Where-Object { $_.LocalName -eq 'Name' }
    if ($attr) {
        $controlName = $attr.Value
        $control = $Form.FindName($controlName)
        if ($control) {
            Set-Variable -Name $controlName -Value $control -Scope Script
        }
    }
}

if ($TxtFooterHint -and $script:LogFilePath) {
    $TxtFooterHint.Text = "Log: $($script:LogFilePath)"
}

# Theme State + Painted Win11 Chrome Hook
$script:IsDarkTheme = Get-SystemIsDarkTheme
$script:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $Form -IsDark $script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $true

$Form.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $Form -IsDark $script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $script:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $script:IsSystemThemeAuto) {
        $Form.Dispatcher.Invoke([Action]{
            $script:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $Form -IsDark $script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $true
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)
$Form.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
})

if ($BtnThemeToggle) {
    $BtnThemeToggle.Add_Click({
        $script:IsSystemThemeAuto = $false
        $script:IsDarkTheme = -not $script:IsDarkTheme
        Set-StandardGuiTheme -Window $Form -IsDark $script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $false
    })
}

# Console/Log Writer helper
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    $uiTimestamp = (Get-Date).ToString("HH:mm:ss")
    $fileTimestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $prefix = "[$uiTimestamp] "
    switch ($Type) {
        "Error"   { $colorPrefix = "[ERROR] ";   $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [ERROR] $Message" }
        "Success" { $colorPrefix = "[SUCCESS] "; $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [SUCCESS] $Message" }
        "Warning" { $colorPrefix = "[WARNING] "; $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [WARNING] $Message" }
        Default   { $colorPrefix = "[INFO] ";    $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [INFO] $Message" }
    }

    if ($script:LogFilePath) {
        try {
            Add-Content -Path $script:LogFilePath -Value $fileMsg -Encoding UTF8 -ErrorAction Stop
        } catch {}
    }

    # Run on UI thread to prevent dispatcher exceptions
    $TxtConsole.Dispatcher.Invoke([Action]{
        $TxtConsole.AppendText("$msg`r`n")
        $TxtConsole.ScrollToEnd()
    })
}

# Reset GUI UI elements to initial state
function Reset-UIState {
    $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushSecBtnBg')
    $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushCardBorder')
    $StatusText.Text = "Disconnected"
    $StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextMuted')

    $BtnDisconnect.Visibility = [System.Windows.Visibility]::Collapsed
    $BtnConnect.Visibility = [System.Windows.Visibility]::Visible
    $BtnConnect.IsEnabled = $true

    $TxtSearch.IsEnabled = $false
    $TxtSearch.Text = ""
    $LstUsers.IsEnabled = $false
    $LstUsers.ItemsSource = $null

    $BtnLoadUsers.IsEnabled = $false

    Reset-SelectedUserDetails
}

# Clear selections and secondary states
function Reset-SelectedUserDetails {
    $TxtSelectedUserName.Text = "No Mailbox Selected"
    $TxtSelectedUserEmail.Text = "Select a mailbox from the left panel to display quarantine status."
    $LstDevices.ItemsSource = $null
    $BtnCheckDevices.IsEnabled = $false
    $BtnRelease.IsEnabled = $false
    $TxtNoDevices.Visibility = [System.Windows.Visibility]::Collapsed
    $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed
}

# Fetch quarantined devices
function Check-QuarantinedDevices {
    param(
        [string]$UserEmail
    )

    if ([string]::IsNullOrWhiteSpace($UserEmail)) { return }

    Write-Log "Checking mobile devices for: $UserEmail"

    # Update UI to loading state
    $LstDevices.ItemsSource = $null
    $TxtNoDevices.Visibility = [System.Windows.Visibility]::Collapsed
    $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Visible
    $BtnCheckDevices.IsEnabled = $false
    $BtnRelease.IsEnabled = $false

    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    # Force UI refresh to display loading state
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        $devices = Get-MobileDevice -Mailbox $UserEmail -ErrorAction Stop

        # Filter only Quarantined
        $quarantined = $devices | Where-Object { $_.DeviceAccessState -eq "Quarantined" }

        $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed

        if ($null -eq $quarantined -or $quarantined.Count -eq 0) {
            Write-Log "No quarantined devices found for $UserEmail." "Info"
            $TxtNoDevices.Visibility = [System.Windows.Visibility]::Visible
            $LstDevices.ItemsSource = $null
            $BtnRelease.IsEnabled = $false
        } else {
            $count = if ($quarantined -is [array]) { $quarantined.Count } else { 1 }
            Write-Log "Found $count quarantined device(s) for $UserEmail." "Warning"

            # Format collection for binding
            $mappedDevices = @()
            foreach ($dev in $quarantined) {
                $mappedDevices += [PSCustomObject]@{
                    DeviceModel       = if ($dev.DeviceModel) { $dev.DeviceModel } else { "Unknown Model" }
                    DeviceType        = if ($dev.DeviceType) { $dev.DeviceType } else { "Unknown Type" }
                    DeviceId          = $dev.DeviceId
                    DeviceAccessState = $dev.DeviceAccessState
                }
            }

            $LstDevices.ItemsSource = $mappedDevices
            $BtnRelease.IsEnabled = $true
        }
    } catch {
        $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed
        Write-Log "Failed to check devices: $_" "Error"
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $BtnCheckDevices.IsEnabled = $true
    }
}

# --- EVENT HANDLERS ---

# Connect to Exchange Online
$BtnConnect.Add_Click({
    Write-Log "Connecting to Exchange Online. Please authenticate via modern login dialog..."
    $BtnConnect.IsEnabled = $false

    # Run Connect
    try {
        # Check and load module if needed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "ExchangeOnlineManagement module not found. Installing module..." "Warning"
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber | Out-Null
            Write-Log "ExchangeOnlineManagement module installed successfully." "Success"
        }

        # Connect to Exchange
        Connect-ExchangeOnline -ErrorAction Stop

        # Resolve administrator identity
        $adminUser = "Exchange Online"
        try {
            $sessions = Get-ConnectionInfo
            if ($sessions -and $sessions.UserName) {
                $adminUser = $sessions.UserName
            }
        } catch {}

        Write-Log "Connection established successfully with $adminUser." "Success"

        # Update Connection Pane UI
        $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushOkTint')
        $StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'BrushOk')
        $StatusText.Text = "Connected ($adminUser)"
        $StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushTextPrimary')

        $BtnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        $BtnConnect.Visibility = [System.Windows.Visibility]::Collapsed
        $BtnLoadUsers.IsEnabled = $true

    } catch {
        Write-Log "Authentication failed: $_" "Error"
        $BtnConnect.IsEnabled = $true
    }
})

# Disconnect Session
$BtnDisconnect.Add_Click({
    Write-Log "Disconnecting active Exchange Online session..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Session disconnected successfully." "Success"
    } catch {
        Write-Log "Error occurred during session teardown: $_" "Warning"
    }
    Reset-UIState
})

# Load User Directory
$BtnLoadUsers.Add_Click({
    Write-Log "Downloading user directory list from tenant..."
    $BtnLoadUsers.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    # Force UI refresh to display loading cursor
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        Write-Log "Retrieving user mailboxes..."
        $script:allUsers = @()

        try {
            $script:allUsers = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, PrimarySmtpAddress |
                               Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName
        } catch {
            Write-Log "Get-EXOMailbox failed. Querying with fallback Get-Mailbox..." "Warning"
            $script:allUsers = Get-Mailbox -ResultSize Unlimited |
                               Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName
        }

        if (-not $script:allUsers) {
            Write-Log "No mailbox users returned from active tenant." "Warning"
            $script:allUsers = @()
        }

        Write-Log "Loaded $($script:allUsers.Count) users successfully." "Success"

        # Bind to UI
        $LstUsers.ItemsSource = $script:allUsers
        $LstUsers.IsEnabled = $true
        $TxtSearch.IsEnabled = $true

    } catch {
        Write-Log "Directory query failed: $_" "Error"
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $BtnLoadUsers.IsEnabled = $true
    }
})

# Filter users directory locally
$TxtSearch.Add_TextChanged({
    if (-not $script:allUsers) { return }
    $filterText = $TxtSearch.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filterText)) {
        $LstUsers.ItemsSource = $script:allUsers
    } else {
        # Filter collection
        $filtered = $script:allUsers | Where-Object {
            $_.DisplayName -like "*$filterText*" -or
            $_.PrimarySmtpAddress -like "*$filterText*" -or
            $_.UserPrincipalName -like "*$filterText*"
        }
        $LstUsers.ItemsSource = $filtered
    }
})

# Selection change event in list of users
$LstUsers.Add_SelectionChanged({
    $selectedUser = $LstUsers.SelectedItem
    if ($null -eq $selectedUser) {
        Reset-SelectedUserDetails
        return
    }

    $TxtSelectedUserName.Text = $selectedUser.DisplayName
    $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
    $TxtSelectedUserEmail.Text = $email

    $BtnCheckDevices.IsEnabled = $true

    # Automatically query mobile devices when user is clicked
    Check-QuarantinedDevices -UserEmail $email
})

# Re-Scan Status Button Click
$BtnCheckDevices.Add_Click({
    $selectedUser = $LstUsers.SelectedItem
    if ($selectedUser) {
        $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
        Check-QuarantinedDevices -UserEmail $email
    }
})

# Release Devices Button Click
$BtnRelease.Add_Click({
    $selectedUser = $LstUsers.SelectedItem
    if (-not $selectedUser) { return }

    $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
    $quarantinedDevices = $LstDevices.ItemsSource

    if (-not $quarantinedDevices) {
        Write-Log "No quarantined devices available to release." "Warning"
        return
    }

    $BtnRelease.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    Write-Log "Executing quarantine release sequence for $email..."

    # Force UI refresh
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        $IDsToRelease = @()
        foreach ($dev in $quarantinedDevices) {
            $IDsToRelease += $dev.DeviceId
        }

        Write-Log "Target Device ID(s): $($IDsToRelease -join ', ')"

        # Fetch existing allowed devices
        Write-Log "Reading current CASMailbox allowed devices..."
        $casMailbox = Get-CASMailbox -Identity $email -ErrorAction Stop
        $CurrentAllowed = $casMailbox.ActiveSyncAllowedDeviceIDs

        # Merge and de-duplicate list
        $UpdatedList = ($CurrentAllowed + $IDsToRelease) | Select-Object -Unique

        # Apply changes
        Write-Log "Writing updated ActiveSyncAllowedDeviceIDs in Exchange Online..."
        Set-CASMailbox -Identity $email -ActiveSyncAllowedDeviceIDs $UpdatedList -ErrorAction Stop

        Write-Log "Quarantine released successfully for $email!" "Success"

        # Recheck devices list to show they are cleared
        Check-QuarantinedDevices -UserEmail $email

    } catch {
        Write-Log "Release sequence failed: $_" "Error"
        $BtnRelease.IsEnabled = $true
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# Teardown Exchange Online session on Close
$Form.Add_Closing({
    Write-Log "Closing Application. Cleaning up..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    } catch {}
})

# Splash Message on Startup
Write-Log "Application initialized. Click 'Connect to Exchange Online' to get started."

# Launch the WPF application
[void]$Form.ShowDialog()
