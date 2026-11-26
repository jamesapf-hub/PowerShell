<#
.SYNOPSIS
    Microsoft Teams Custom Background Win32 App Packager for Intune
.DESCRIPTION
    A modern WPF GUI tool that packages custom Microsoft Teams background images
    into an Intune Win32 application (.intunewin). Generates automated deployment,
    uninstallation, and detection scripts, downloads IntuneWinAppUtil if needed,
    and outputs all exact Intune configuration settings.
.EXAMPLE
    .\TeamsBackgroundPackager.ps1
.NOTES
    Script Name   : TeamsBackgroundPackager.ps1
    Author        : JP
    Version       : 1.2.0
    Created       : 10/08/26
    Last Modified : 16/11/26
    Log Path      : C:\Logs\TeamsBackgrounds\
#>

# Ensure Single-Threaded Apartment (STA) mode required for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) { $scriptPath = $PSCommandPath }
    if ($scriptPath) {
        Write-Host "Re-launching in STA mode..." -ForegroundColor Cyan
        Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
        exit
    }
}

# Add required assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Windows.Forms, System.Xaml

# Setup working directories
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$ToolsDir = Join-Path $ScriptDir "tools"
$DefaultOutputDir = Join-Path $ScriptDir "Output"

if (-not (Test-Path $ToolsDir)) {
    New-Item -Path $ToolsDir -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $DefaultOutputDir)) {
    New-Item -Path $DefaultOutputDir -ItemType Directory -Force | Out-Null
}

# Image Items Collection (ObservableCollection for UI data binding)
$Global:ImageItems = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

# Class to store image item details (compiled only once per session)
if (-not ([System.Management.Automation.PSTypeName]'TeamsImageItem').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Windows.Media.Imaging;

public class TeamsImageItem
{
    public string FileName { get; set; }
    public string FullPath { get; set; }
    public string Resolution { get; set; }
    public string FileSize { get; set; }
    public string AspectRatio { get; set; }
    public bool IsRecommendedSize { get; set; }
    public BitmapImage Thumbnail { get; set; }
}
"@ -ReferencedAssemblies "PresentationCore", "WindowsBase", "PresentationFramework", "System.Xaml"
}

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
            BrushOk            = '#34D399'
            BrushErr           = '#F87171'
            BrushDangerBg      = '#26EF4444'
            BrushDangerHover   = '#45EF4444'
            BrushDangerPressed = '#66EF4444'
            BrushDangerBorder  = '#66EF4444'
            BrushDangerFg      = '#FCA5A5'
            BrushDangerHoverFg = '#FFFFFF'
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
            BrushOk            = '#059669'
            BrushErr           = '#DC2626'
            BrushDangerBg      = '#FEE2E2'
            BrushDangerHover   = '#FECACA'
            BrushDangerPressed = '#FCA5A5'
            BrushDangerBorder  = '#EF4444'
            BrushDangerFg      = '#B91C1C'
            BrushDangerHoverFg = '#7F1D1D'
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
        [System.Windows.Window]$Window
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $resolvedFontDir = $null

    $searchRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        Split-Path -Path $PSCommandPath -Parent
    } else {
        $null
    }

    if ($searchRoot) {
        $probe = $searchRoot
        for ($i = 0; $i -lt 5 -and -not [string]::IsNullOrWhiteSpace($probe); $i++) {
            $candidateDir = Join-Path -Path $probe -ChildPath 'assets\fonts'
            if (Test-Path -Path (Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                $resolvedFontDir = $candidateDir
                break
            }
            $probe = Split-Path -Path $probe -Parent
        }
    }

    if (-not $resolvedFontDir) {
        $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        try {
            if (-not (Test-Path -Path $cacheDir)) {
                New-Item -Path $cacheDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            $allCached = $true
            foreach ($ff in $fontFiles) {
                if (-not (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath $ff) -PathType Leaf)) {
                    $allCached = $false
                    break
                }
            }
            if (-not $allCached) {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    foreach ($ff in $fontFiles) {
                        $dest = Join-Path -Path $cacheDir -ChildPath $ff
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            $url = "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/$ff"
                            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
                        }
                    }
                } finally {
                    $ProgressPreference = $prevProgress
                }
            }
            if (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                $resolvedFontDir = $cacheDir
            }
        } catch {
            Write-Verbose "Figtree font download skipped: $($_.Exception.Message)"
        }
    }

    if ($resolvedFontDir) {
        try {
            $fontUri = ([System.Uri]($resolvedFontDir.TrimEnd('\') + '\')).AbsoluteUri
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
    $Window.Left   = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $Window.Top    = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Window.MinWidth  = [Math]::Min(860, [int]$targetWidth)
    $Window.MinHeight = [Math]::Min(640, [int]$targetHeight)
    $Window.ResizeMode = 'CanResize'
}

# XAML Definition
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Teams Background Packager v1.2.0"
        Height="820" Width="1080" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual"
        Background="#091122"
        Foreground="{DynamicResource BrushTextPrimary}">

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
        <SolidColorBrush x:Key="BrushOk" Color="#34D399"/>
        <SolidColorBrush x:Key="BrushErr" Color="#F87171"/>
        <SolidColorBrush x:Key="BrushDangerBg" Color="#26EF4444"/>
        <SolidColorBrush x:Key="BrushDangerHover" Color="#45EF4444"/>
        <SolidColorBrush x:Key="BrushDangerPressed" Color="#66EF4444"/>
        <SolidColorBrush x:Key="BrushDangerBorder" Color="#66EF4444"/>
        <SolidColorBrush x:Key="BrushDangerFg" Color="#FCA5A5"/>
        <SolidColorBrush x:Key="BrushDangerHoverFg" Color="#FFFFFF"/>

        <!-- Slim ScrollBar Style -->
        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="8"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="Transparent">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border Background="{DynamicResource BrushScrollThumb}" CornerRadius="4" Margin="1"/>
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

        <!-- Base Secondary Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushSecBtnPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary CTA Button -->
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary Build CTA Button -->
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="FontWeight" Value="Bold"/>
        </Style>

        <!-- Danger Button -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{DynamicResource BrushDangerBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushDangerBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushDangerFg}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushDangerHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource BrushDangerPressed}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="{DynamicResource BrushDangerHoverFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Secondary Button -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>

        <!-- Styled TextBox -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="0,3"/>
        </Style>

        <!-- RadioButton Style -->
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="0,3"/>
        </Style>

        <!-- Surface Panel Style -->
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource BrushCardBg}"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>

        <!-- GridView Header & Item Styles -->
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontSize" Value="11.5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>

        <Style TargetType="ListViewItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="4,4"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BrushAccentTint}"/>
                    <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Banner -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Teams Packager Icon -->
                <Border Grid.Column="0" Background="{DynamicResource BrushAccentTint}" BorderBrush="{DynamicResource BrushAccent}" BorderThickness="1" Width="44" Height="44" CornerRadius="8" Margin="0,0,14,0">
                    <Viewbox Width="24" Height="24" HorizontalAlignment="Center" VerticalAlignment="Center">
                        <Canvas Width="24" Height="24">
                            <Path Fill="{DynamicResource BrushAccent}" Data="M16.5 12c1.38 0 2.49-1.12 2.49-2.5S17.88 7 16.5 7 14 8.12 14 9.5s1.11 2.5 2.5 2.5zm-9 0c1.38 0 2.49-1.12 2.49-2.5S8.88 7 7.5 7 5 8.12 5 9.5 6.11 12 7.5 12zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-1.5c0-2.33-4.67-3.5-7-3.5zm9 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-1.5c0-2.33-4.67-3.5-7-3.5z"/>
                        </Canvas>
                    </Viewbox>
                </Border>

                <!-- Titles -->
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="Microsoft Teams Background Packager" FontSize="20" FontWeight="Bold" Foreground="{DynamicResource BrushTextPrimary}"/>
                    <TextBlock Text="Package custom video meeting backgrounds into a production-ready Intune Win32 (.intunewin) application" FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}" Margin="0,2,0,0"/>
                </StackPanel>

                <!-- Theme Toggle & Version Badge -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnThemeToggle" Padding="12,6" Margin="0,0,10,0" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock x:Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <Border Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="12" Padding="12,4" VerticalAlignment="Center">
                        <TextBlock Text="v1.2.0 | Win32 App" FontSize="11.5" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main Tab Control -->
        <TabControl Grid.Row="1" x:Name="MainTabs" Background="Transparent" BorderThickness="0" Padding="0" Margin="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border x:Name="TabBorder" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2" Padding="18,10" Margin="0,0,6,0">
                                    <ContentPresenter ContentSource="Header" TextBlock.FontSize="13.5" TextBlock.FontWeight="SemiBold"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource BrushCardBg}"/>
                                        <Setter TargetName="TabBorder" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                                        <Setter Property="Foreground" Value="{DynamicResource BrushAccent}"/>
                                    </Trigger>
                                    <Trigger Property="IsSelected" Value="False">
                                        <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </TabControl.Resources>

            <!-- TAB 1: BUILD PACKAGE -->
            <TabItem Header="Package Builder">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="390"/>
                        <ColumnDefinition Width="1"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left Column: Settings -->
                    <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="20,16">
                        <StackPanel>
                            <TextBlock Text="PACKAGE CONFIGURATION" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,0,0,10"/>

                            <TextBlock Text="Application Name" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,0,0,4"/>
                            <TextBox x:Name="txtAppName" Text="Teams Backgrounds - Corporate Branding" Margin="0,0,0,10"/>

                            <TextBlock Text="Publisher" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,0,0,4"/>
                            <TextBox x:Name="txtPublisher" Text="Corporate IT" Margin="0,0,0,10"/>

                            <TextBlock Text="Version" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,0,0,4"/>
                            <TextBox x:Name="txtVersion" Text="1.0.0" Margin="0,0,0,10"/>

                            <TextBlock Text="Description" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,0,0,4"/>
                            <TextBox x:Name="txtDescription" Text="Standard corporate branded background images for Microsoft Teams video meetings." TextWrapping="Wrap" AcceptsReturn="True" Height="50" Margin="0,0,0,12"/>

                            <TextBlock Text="TARGET TEAMS CLIENTS" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,6,0,6"/>
                            <CheckBox x:Name="chkNewTeams" Content="New Microsoft Teams (MSTeams / v2)" IsChecked="True"/>
                            <CheckBox x:Name="chkClassicTeams" Content="Classic Microsoft Teams (v1)" IsChecked="True" Margin="0,2,0,12"/>

                            <TextBlock Text="DEPLOYMENT CONTEXT" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,6,0,6"/>
                            <RadioButton x:Name="rbAutoContext" GroupName="Context" Content="Smart Auto-Detect (Recommended)" IsChecked="True"/>
                            <TextBlock Text="Deploys to all user profiles in C:\Users (including Sandbox &amp; Default profile) when running as SYSTEM. Deploys to current user if run as User." FontSize="11" Foreground="{DynamicResource BrushTextMuted}" Margin="20,0,0,8" TextWrapping="Wrap"/>

                            <RadioButton x:Name="rbUserContext" GroupName="Context" Content="User Context"/>
                            <TextBlock Text="Configured for Intune 'User' context (with fail-safe fallback if run as SYSTEM)." FontSize="11" Foreground="{DynamicResource BrushTextMuted}" Margin="20,0,0,8" TextWrapping="Wrap"/>

                            <RadioButton x:Name="rbSystemContext" GroupName="Context" Content="System Context (All Users)"/>
                            <TextBlock Text="Always installs to all user profiles under C:\Users (Intune: 'System')." FontSize="11" Foreground="{DynamicResource BrushTextMuted}" Margin="20,0,0,12" TextWrapping="Wrap"/>

                            <TextBlock Text="LOG DIRECTORY (ENDPOINTS)" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,6,0,4"/>
                            <TextBox x:Name="txtLogDir" Text="C:\Logs\TeamsBackgrounds" Margin="0,0,0,12"/>

                            <TextBlock Text="OUTPUT DIRECTORY" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,6,0,4"/>
                            <Grid Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtOutputDir" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0"/>
                                <Button x:Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Padding="12,6"/>
                            </Grid>
                        </StackPanel>
                    </ScrollViewer>

                    <!-- 1px Hairline Divider -->
                    <Border Grid.Column="1" Background="{DynamicResource BrushHairline}"/>

                    <!-- Right Column: Images List -->
                    <Grid Grid.Column="2" Margin="18,14,20,14">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Action Toolbar -->
                        <Border Grid.Row="0" Style="{StaticResource CardBorder}" Padding="12,10" Margin="0,0,0,10">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0" Orientation="Horizontal">
                                    <Button x:Name="btnAddImages" Content="+ Add Images..." Style="{StaticResource PrimaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnRemoveSelected" Content="Remove Selected" Style="{StaticResource DangerButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnClearAll" Content="Clear All" Margin="0,0,8,0"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="Backgrounds: " FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}"/>
                                    <TextBlock x:Name="lblImageCount" Text="0 images" FontSize="12.5" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Images ListView -->
                        <Border Grid.Row="1" Background="{DynamicResource BrushCardBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="8" Padding="4">
                            <Grid>
                                <ListView x:Name="lvImages" Background="Transparent" BorderThickness="0" Foreground="{DynamicResource BrushTextPrimary}"
                                          SelectionMode="Extended" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                                    <ListView.View>
                                        <GridView>
                                            <GridViewColumn Header="Preview" Width="110">
                                                <GridViewColumn.CellTemplate>
                                                    <DataTemplate>
                                                        <Image Source="{Binding Thumbnail}" Width="96" Height="54" Stretch="UniformToFill" Margin="2"/>
                                                    </DataTemplate>
                                                </GridViewColumn.CellTemplate>
                                            </GridViewColumn>
                                            <GridViewColumn Header="File Name" Width="190" DisplayMemberBinding="{Binding FileName}"/>
                                            <GridViewColumn Header="Resolution" Width="110" DisplayMemberBinding="{Binding Resolution}"/>
                                            <GridViewColumn Header="Size" Width="80" DisplayMemberBinding="{Binding FileSize}"/>
                                            <GridViewColumn Header="Status" Width="110" DisplayMemberBinding="{Binding AspectRatio}"/>
                                            <GridViewColumn Header="Source Path" Width="200" DisplayMemberBinding="{Binding FullPath}"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>

                                <!-- Empty State Placeholder -->
                                <StackPanel x:Name="pnlEmptyState" VerticalAlignment="Center" HorizontalAlignment="Center">
                                    <Viewbox Width="54" Height="54" HorizontalAlignment="Center" Margin="0,0,0,12">
                                        <Canvas Width="24" Height="24">
                                            <Path Fill="{DynamicResource BrushAccent}" Opacity="0.45" Data="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/>
                                        </Canvas>
                                    </Viewbox>
                                    <TextBlock Text="No Teams background images added yet" FontSize="15" FontWeight="SemiBold" HorizontalAlignment="Center" Foreground="{DynamicResource BrushTextSecondary}"/>
                                    <TextBlock Text="Click 'Add Images' to select .PNG or .JPG corporate backgrounds (1920x1080 recommended)" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" HorizontalAlignment="Center" Margin="0,6,0,16"/>
                                    <Button x:Name="btnAddImagesEmpty" Content="+ Select Images" Style="{StaticResource PrimaryButton}" HorizontalAlignment="Center" Padding="20,10"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: INTUNE GUIDANCE -->
            <TabItem x:Name="tabIntuneGuidance" Header="Intune Deployment Details">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="20,14">
                    <StackPanel>
                        <!-- Package Status Banner -->
                        <Border Style="{StaticResource CardBorder}" BorderBrush="{DynamicResource BrushAccent}">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Viewbox Grid.Column="0" Width="32" Height="32" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,0,14,0">
                                    <Canvas Width="24" Height="24">
                                        <Path Fill="{DynamicResource BrushOk}" Data="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                                    </Canvas>
                                </Viewbox>
                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                    <TextBlock Text="Win32 App Package Built Successfully" FontSize="15.5" FontWeight="Bold" Foreground="{DynamicResource BrushTextPrimary}"/>
                                    <TextBlock x:Name="lblPackagePath" Text="Package path will appear here once built." FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,3,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <Button x:Name="btnOpenOutputDir" Content="Open Folder" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyPackagePath" Content="Copy Path" Style="{StaticResource PrimaryButton}"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Step 1: Program Settings -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="STEP 1: INTUNE PROGRAM SETTINGS" FontSize="11.5" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,0,0,12"/>

                                <Grid Margin="0,0,0,12">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Install Command" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtInstallCommand" Grid.Row="1" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyInstallCmd" Grid.Row="1" Grid.Column="1" Content="Copy" Padding="14,6"/>
                                </Grid>

                                <Grid Margin="0,0,0,12">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Uninstall Command" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtUninstallCommand" Grid.Row="1" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyUninstallCmd" Grid.Row="1" Grid.Column="1" Content="Copy" Padding="14,6"/>
                                </Grid>

                                <Grid Margin="0,0,0,4">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>

                                    <StackPanel Grid.Column="0" Margin="0,0,8,0">
                                        <TextBlock Text="Install Behavior" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                                        <TextBox x:Name="txtInstallBehavior" Text="User" IsReadOnly="True"/>
                                    </StackPanel>

                                    <StackPanel Grid.Column="1" Margin="0,0,8,0">
                                        <TextBlock Text="Device Restart Behavior" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                                        <TextBox Text="Determine behavior based on return codes" IsReadOnly="True"/>
                                    </StackPanel>

                                    <StackPanel Grid.Column="2">
                                        <TextBlock Text="Return Code for Success" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,4"/>
                                        <TextBox Text="0 (Success)" IsReadOnly="True"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Step 2: Detection Rules -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Text="STEP 2: DETECTION RULES CONFIGURATION" FontSize="11.5" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}"/>
                                        <TextBlock Text="Rules format: 'Use a custom detection script'. Select the generated Detect-TeamsBackgrounds.ps1 file." FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,2,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                                        <Button x:Name="btnCopyDetectionScript" Content="Copy Script" Margin="0,0,8,0"/>
                                        <Button x:Name="btnOpenDetectionFile" Content="Open Script File"/>
                                    </StackPanel>
                                </Grid>

                                <TextBox x:Name="txtDetectionScriptPreview" IsReadOnly="True" AcceptsReturn="True"
                                         FontFamily="Cascadia Mono, Consolas" FontSize="12" Background="{DynamicResource BrushLogBg}"
                                         BorderBrush="{DynamicResource BrushCardBorder}" Foreground="{DynamicResource BrushTextSecondary}"
                                         Height="180" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                         Padding="10"/>
                            </StackPanel>
                        </Border>

                        <!-- Step 3: Requirements Guide -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="STEP 3: REQUIREMENTS &amp; ASSIGNMENT" FontSize="11.5" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}" Margin="0,0,0,8"/>
                                <TextBlock Text="- Operating System Architecture: 64-bit and 32-bit" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,2"/>
                                <TextBlock Text="- Minimum Operating System: Windows 10 1909 or higher" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,2"/>
                                <TextBlock Text="- Assignment: Assign as 'Required' to targeted Entra ID User / Device groups." FontSize="12" Foreground="{DynamicResource BrushTextSecondary}" Margin="0,2"/>
                            </StackPanel>
                        </Border>

                        <!-- Step 4: Enterprise Logging Guide -->
                        <Border Style="{StaticResource CardBorder}" Margin="0">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Text="STEP 4: LOGGING &amp; DIAGNOSTICS" FontSize="11.5" FontWeight="Bold" Foreground="{DynamicResource BrushAccent}"/>
                                        <TextBlock Text="Full runtime diagnostic logs are automatically written to this folder on target machines:" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" Margin="0,2,0,0"/>
                                    </StackPanel>
                                    <Button x:Name="btnOpenLogDir" Grid.Column="1" Content="Open Log Folder" Padding="12,6"/>
                                </Grid>

                                <Grid Margin="0,4,0,3">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Log Directory:" FontSize="12" Foreground="{DynamicResource BrushTextMuted}"/>
                                    <TextBlock x:Name="lblLogDirPath" Grid.Column="1" Text="C:\Logs\TeamsBackgrounds" FontSize="12" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Install Log:" FontSize="12" Foreground="{DynamicResource BrushTextMuted}"/>
                                    <TextBlock x:Name="lblInstallLogPath" Grid.Column="1" Text="C:\Logs\TeamsBackgrounds\Install.log" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Detection Log:" FontSize="12" Foreground="{DynamicResource BrushTextMuted}"/>
                                    <TextBlock x:Name="lblDetectLogPath" Grid.Column="1" Text="C:\Logs\TeamsBackgrounds\Detection.log" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Uninstall Log:" FontSize="12" Foreground="{DynamicResource BrushTextMuted}"/>
                                    <TextBlock x:Name="lblUninstallLogPath" Grid.Column="1" Text="C:\Logs\TeamsBackgrounds\Uninstall.log" FontSize="12" Foreground="{DynamicResource BrushTextSecondary}"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Bottom Action Bar -->
        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Status & Progress -->
                <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,20,0">
                    <TextBlock x:Name="lblStatus" Text="Ready. Add background images and click 'Build .intunewin Package'." FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}" TextTrimming="CharacterEllipsis"/>
                    <ProgressBar x:Name="progressBar" Height="5" Margin="0,6,0,0" Background="{DynamicResource BrushCardBorder}" Foreground="{DynamicResource BrushAccent}" BorderThickness="0" Value="0" Visibility="Collapsed"/>
                </StackPanel>

                <!-- Build Button -->
                <Button x:Name="btnBuildPackage" Grid.Column="1" Content="Build .intunewin Package" Style="{StaticResource SuccessButton}" FontSize="13.5" Padding="22,9"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Helper to load XAML cleanly
[xml]$xml = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xml
$Window = [System.Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $Window

# Map UI Controls
$btnThemeToggle = $Window.FindName("BtnThemeToggle")
$txtThemeLabel = $Window.FindName("TxtThemeLabel")
$txtAppName = $Window.FindName("txtAppName")
$txtPublisher = $Window.FindName("txtPublisher")
$txtVersion = $Window.FindName("txtVersion")
$txtDescription = $Window.FindName("txtDescription")
$chkNewTeams = $Window.FindName("chkNewTeams")
$chkClassicTeams = $Window.FindName("chkClassicTeams")
$rbAutoContext = $Window.FindName("rbAutoContext")
$rbUserContext = $Window.FindName("rbUserContext")
$rbSystemContext = $Window.FindName("rbSystemContext")
$txtOutputDir = $Window.FindName("txtOutputDir")
$btnBrowseOutput = $Window.FindName("btnBrowseOutput")

$btnAddImages = $Window.FindName("btnAddImages")
$btnAddImagesEmpty = $Window.FindName("btnAddImagesEmpty")
$btnRemoveSelected = $Window.FindName("btnRemoveSelected")
$btnClearAll = $Window.FindName("btnClearAll")
$lblImageCount = $Window.FindName("lblImageCount")
$lvImages = $Window.FindName("lvImages")
$pnlEmptyState = $Window.FindName("pnlEmptyState")

$MainTabs = $Window.FindName("MainTabs")
$tabIntuneGuidance = $Window.FindName("tabIntuneGuidance")
$lblPackagePath = $Window.FindName("lblPackagePath")
$btnOpenOutputDir = $Window.FindName("btnOpenOutputDir")
$btnCopyPackagePath = $Window.FindName("btnCopyPackagePath")

$txtInstallCommand = $Window.FindName("txtInstallCommand")
$txtUninstallCommand = $Window.FindName("txtUninstallCommand")
$btnCopyInstallCmd = $Window.FindName("btnCopyInstallCmd")
$btnCopyUninstallCmd = $Window.FindName("btnCopyUninstallCmd")
$txtInstallBehavior = $Window.FindName("txtInstallBehavior")

$txtDetectionScriptPreview = $Window.FindName("txtDetectionScriptPreview")
$btnCopyDetectionScript = $Window.FindName("btnCopyDetectionScript")
$btnOpenDetectionFile = $Window.FindName("btnOpenDetectionFile")

$lblStatus = $Window.FindName("lblStatus")
$progressBar = $Window.FindName("progressBar")
$btnBuildPackage = $Window.FindName("btnBuildPackage")

$txtLogDir = $Window.FindName("txtLogDir")
$btnOpenLogDir = $Window.FindName("btnOpenLogDir")
$lblLogDirPath = $Window.FindName("lblLogDirPath")
$lblInstallLogPath = $Window.FindName("lblInstallLogPath")
$lblDetectLogPath = $Window.FindName("lblDetectLogPath")
$lblUninstallLogPath = $Window.FindName("lblUninstallLogPath")

# Initialize default values
$txtOutputDir.Text = $DefaultOutputDir
$lvImages.ItemsSource = $Global:ImageItems
$Global:LastBuiltDetectionScriptPath = ""
$Global:LastBuiltPackagePath = ""

$Script:IsDarkTheme = Get-SystemIsDarkTheme
$Script:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true

$Window.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $Script:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $Script:IsSystemThemeAuto) {
        $Window.Dispatcher.Invoke([Action]{
            $Script:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)
$Window.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
})

$btnThemeToggle.Add_Click({
    $Script:IsSystemThemeAuto = $false
    $Script:IsDarkTheme = -not $Script:IsDarkTheme
    Set-StandardGuiTheme -Window $Window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $false
})

# Function to update UI image list counter & empty state
function Update-ImageCounter {
    $count = $Global:ImageItems.Count
    if ($count -eq 1) {
        $lblImageCount.Text = "1 image"
    } else {
        $lblImageCount.Text = "$count images"
    }

    if ($count -eq 0) {
        $pnlEmptyState.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $pnlEmptyState.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Function to add image files to collection
function Add-ImageFile([string[]]$filePaths) {
    foreach ($path in $filePaths) {
        if (-not (Test-Path $path)) { continue }
        
        $fileName = [System.IO.Path]::GetFileName($path)
        
        # Check duplicate
        $existing = $Global:ImageItems | Where-Object { $_.FileName -eq $fileName }
        if ($existing) {
            continue
        }

        try {
            $fileInfo = New-Object System.IO.FileInfo($path)
            $sizeStr = ""
            if ($fileInfo.Length -gt 1MB) {
                $sizeStr = "{0:N2} MB" -f ($fileInfo.Length / 1MB)
            } else {
                $sizeStr = "{0:N0} KB" -f ($fileInfo.Length / 1KB)
            }

            # Extract dimensions using System.Drawing
            $img = [System.Drawing.Image]::FromFile($path)
            $w = $img.Width
            $h = $img.Height
            $img.Dispose()

            $resStr = "${w}x${h}"
            $isOptimal = ($w -eq 1920 -and $h -eq 1080)
            $aspectStr = if ($isOptimal) { "Optimal (16:9)" } else { "Custom (${w}x${h})" }

            # Create BitmapImage with CacheOption OnLoad so file isn't locked
            $bi = New-Object System.Windows.Media.Imaging.BitmapImage
            $bi.BeginInit()
            $bi.UriSource = New-Object System.Uri($path, [System.UriKind]::Absolute)
            $bi.DecodePixelWidth = 120
            $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bi.EndInit()
            $bi.Freeze()

            $item = New-Object TeamsImageItem
            $item.FileName = $fileName
            $item.FullPath = $path
            $item.Resolution = $resStr
            $item.FileSize = $sizeStr
            $item.AspectRatio = $aspectStr
            $item.IsRecommendedSize = $isOptimal
            $item.Thumbnail = $bi

            $Global:ImageItems.Add($item)
        } catch {
            [System.Windows.MessageBox]::Show("Failed to load image '$fileName': $_", "Image Load Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
    }
    Update-ImageCounter
}

# Image selection dialog
$SelectImagesAction = {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select Teams Background Images"
    $dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|PNG Images (*.png)|*.png|JPEG Images (*.jpg;*.jpeg)|*.jpg;*.jpeg|All Files (*.*)|*.*"
    $dlg.Multiselect = $true
    
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-ImageFile $dlg.FileNames
    }
}

$btnAddImages.Add_Click($SelectImagesAction)
$btnAddImagesEmpty.Add_Click($SelectImagesAction)

# Remove selected images
$btnRemoveSelected.Add_Click({
    $selected = @($lvImages.SelectedItems)
    if ($selected.Count -eq 0) { return }
    foreach ($item in $selected) {
        $Global:ImageItems.Remove($item)
    }
    Update-ImageCounter
})

# Clear all images
$btnClearAll.Add_Click({
    if ($Global:ImageItems.Count -eq 0) { return }
    $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove all images from the list?", "Confirm Clear", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $Global:ImageItems.Clear()
        Update-ImageCounter
    }
})

# Browse output directory
$btnBrowseOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Output Directory for .intunewin Package"
    $dlg.SelectedPath = $txtOutputDir.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $dlg.SelectedPath
    }
})

# Context radio button changes update guidance display
$rbAutoContext.Add_Checked({
    $txtInstallBehavior.Text = "System (or User)"
})
$rbUserContext.Add_Checked({
    $txtInstallBehavior.Text = "User"
})
$rbSystemContext.Add_Checked({
    $txtInstallBehavior.Text = "System"
})

# Function to ensure IntuneWinAppUtil.exe is available
function Get-IntuneWinAppUtilPath {
    $expectedPath = Join-Path $ToolsDir "IntuneWinAppUtil.exe"
    if (Test-Path $expectedPath) {
        return $expectedPath
    }

    # Check PATH
    $cmd = try { Get-Command "IntuneWinAppUtil.exe" -ErrorAction Stop } catch { $null }
    if ($cmd) {
        return $cmd.Source
    }

    # Attempt download from official Microsoft GitHub repository
    $lblStatus.Text = "Downloading official IntuneWinAppUtil.exe from Microsoft..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $url = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $expectedPath -UseBasicParsing -TimeoutSec 30
        $sig = Get-AuthenticodeSignature -FilePath $expectedPath -ErrorAction Stop
        if ($sig.Status -ne 'Valid') {
            try { Remove-Item $expectedPath -Force -ErrorAction Stop } catch {}
            throw "Authenticode signature check failed for IntuneWinAppUtil.exe ($($sig.Status))"
        }
        if (Test-Path $expectedPath) {
            return $expectedPath
        }
    } catch {
        Write-Warning "Auto-download failed: $_"
    }

    # If download fails, prompt user to locate it
    $result = [System.Windows.MessageBox]::Show("IntuneWinAppUtil.exe was not found. Would you like to browse and locate it manually?", "IntuneWinAppUtil Required", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "Locate IntuneWinAppUtil.exe"
        $dlg.Filter = "Executable Files (*.exe)|*.exe"
        $dlg.FileName = "IntuneWinAppUtil.exe"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Copy-Item -Path $dlg.FileName -Destination $expectedPath -Force
            return $expectedPath
        }
    }

    return $null
}

# Function to generate the Install script
function New-InstallScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$Version, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Installs corporate Microsoft Teams background images.
    Auto-generated by Teams Background Packager for Intune.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$AppName = "__APPNAME__"
$AppVersion = "__APPVERSION__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__" # "Auto", "User", "System"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ImageFiles = @(__IMAGEFILES__)

# Initialize Enterprise Log Directory (Primary: C:\Log\<AppName>, Fallback: %LOCALAPPDATA%\Log\<AppName>)
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    # Test write access
    $testFile = Join-Path $LogDir ".write_test"
    [System.IO.File]::WriteAllText($testFile, "test")
    try { Remove-Item -Path $testFile -Force -ErrorAction Stop } catch {}

    # If running elevated or as SYSTEM, ensure Users have modify access to C:\Logs
    $isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isElevated -and (Test-Path "C:\Logs")) {
        try {
            $acl = Get-Acl "C:\Logs"
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Users", "Modify,Synchronize", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($rule)
            Set-Acl -Path "C:\Logs" -AclObject $acl -ErrorAction Stop
        } catch {}
    }
} catch {
    $fallbackDir = "$env:ProgramData\Logs\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        try { New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {}
    }
}

$LogFile = Join-Path $LogDir "Install.log"

function Write-Log([string]$message) {
    $stamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $entry = "[$stamp] $message"
    Write-Output $entry
    try { Add-Content -Path $LogFile -Value $entry -ErrorAction Stop } catch {}
}

# Determine runtime identity
$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

# Auto-detect or fail-safe:
# If running as SYSTEM (Sandbox runner, Intune System Context, SCCM), NEVER deploy to systemprofile!
# Always deploy across all real user profiles in C:\Users (including WDAGUtilityAccount in Sandbox & Default)
if ($IsSystemIdentity) {
    $EffectiveMode = "AllProfiles"
    $ContextLabel = "SYSTEM Context (Multi-Profile Auto-Detect)"
} elseif ($ConfiguredMode -eq "System") {
    $EffectiveMode = "AllProfiles"
    $ContextLabel = "System Context (All Profiles Configured)"
} else {
    $EffectiveMode = "CurrentUser"
    $ContextLabel = "User Context ($env:USERNAME)"
}

Write-Log "========================================================"
Write-Log "Starting installation of $AppName (v$AppVersion)..."
Write-Log "Host Computer:      $env:COMPUTERNAME"
Write-Log "Executing Identity: $($CurrentIdentity.Name)"
Write-Log "Execution Context:  $ContextLabel"
if ($IsSystemIdentity -and ($ConfiguredMode -eq "User")) {
    Write-Log "NOTICE: Running as SYSTEM/Sandbox. Auto-routing deployment to all user profiles in C:\Users (including WDAGUtilityAccount & Default) so real users receive backgrounds."
}
Write-Log "Active Log File:    $LogFile"
Write-Log "Background Count:   $($ImageFiles.Count) images"
Write-Log "========================================================"

$SourceImagesFolder = Join-Path $PSScriptRoot "Images"
if (-not (Test-Path $SourceImagesFolder)) {
    Write-Log "ERROR: Source Images folder not found at $SourceImagesFolder"
    exit 1
}

function Deploy-ToFolder([string]$targetFolder, [string]$folderLabel) {
    if (-not (Test-Path $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created directory ($folderLabel): $targetFolder"
    }

    foreach ($img in $ImageFiles) {
        $sourcePath = Join-Path $SourceImagesFolder $img
        $destPath = Join-Path $targetFolder $img
        if (Test-Path $sourcePath) {
            $fInfo = Get-Item $sourcePath
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            $sizeKb = [Math]::Round($fInfo.Length / 1KB, 1)
            Write-Log "Copied ($folderLabel): '$img' ($sizeKb KB) -> '$targetFolder'"
        } else {
            Write-Log "WARNING: Source image '$sourcePath' does not exist."
        }
    }
}

try {
    if ($EffectiveMode -eq "CurrentUser") {
        # Current User Context
        Write-Log "Deploying in User Context for: $env:USERNAME"

        if ($TargetNewTeams) {
            $newTeamsUploads = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Deploy-ToFolder -targetFolder $newTeamsUploads -folderLabel "New Teams"
        }

        if ($TargetClassicTeams) {
            $classicTeamsUploads = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
            Deploy-ToFolder -targetFolder $classicTeamsUploads -folderLabel "Classic Teams"
        }

        # Write User marker file
        $markerDir = "$env:LOCALAPPDATA\TeamsCustomBackgrounds"
        if (-not (Test-Path $markerDir)) { New-Item -Path $markerDir -ItemType Directory -Force | Out-Null }
        $manifest = @{
            AppName       = $AppName
            Version       = $AppVersion
            InstalledDate = (Get-Date).ToString("o")
            Context       = "User"
            Images        = $ImageFiles
            LogDir        = $LogDir
        }
        $manifest | ConvertTo-Json | Set-Content -Path (Join-Path $markerDir "installed.json") -Force
        Write-Log "Wrote user marker file to $markerDir\installed.json"

    } else {
        # Deploy across all user profiles in C:\Users (including Sandbox accounts like WDAGUtilityAccount and Default)
        Write-Log "Deploying across user profiles in C:\Users..."

        $userProfiles = Get-ChildItem -Path "C:\Users" -Directory -Force | Where-Object {
            $_.Name -notmatch '^(Public|All Users|Default User)$' -and
            -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
            ($_.Name -notmatch '^[\!\(\}]')
        }

        foreach ($profile in $userProfiles) {
            Write-Log "Processing profile: $($profile.FullName)"

            if ($TargetNewTeams) {
                $userNewTeams = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
                Deploy-ToFolder -targetFolder $userNewTeams -folderLabel "Profile: $($profile.Name) - New Teams"
            }

            if ($TargetClassicTeams) {
                $userClassicTeams = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
                Deploy-ToFolder -targetFolder $userClassicTeams -folderLabel "Profile: $($profile.Name) - Classic Teams"
            }

            # Drop user marker inside profile
            $userMarkerDir = Join-Path $profile.FullName "AppData\Local\TeamsCustomBackgrounds"
            if (-not (Test-Path $userMarkerDir)) { try { New-Item -Path $userMarkerDir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {} }
            $userManifest = @{
                AppName       = $AppName
                Version       = $AppVersion
                InstalledDate = (Get-Date).ToString("o")
                Context       = "System-Deployed"
                Images        = $ImageFiles
            }
            try { $userManifest | ConvertTo-Json | Set-Content -Path (Join-Path $userMarkerDir "installed.json") -Force -ErrorAction Stop } catch {}
        }

        # Write System marker file
        $markerDir = "$env:ProgramData\TeamsCustomBackgrounds"
        if (-not (Test-Path $markerDir)) { New-Item -Path $markerDir -ItemType Directory -Force | Out-Null }
        $manifest = @{
            AppName       = $AppName
            Version       = $AppVersion
            InstalledDate = (Get-Date).ToString("o")
            Context       = "System"
            Images        = $ImageFiles
            LogDir        = $LogDir
        }
        $manifest | ConvertTo-Json | Set-Content -Path (Join-Path $markerDir "installed.json") -Force
        Write-Log "Wrote system marker file to $markerDir\installed.json"
    }

    Write-Log "Installation completed successfully."
    exit 0
} catch {
    Write-Log "ERROR during installation: $_"
    exit 1
}
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__APPVERSION__", $Version)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Function to generate the Detection script
function New-DetectScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$Version, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Detection script for corporate Microsoft Teams background images.
    Returns 0 and outputs success text if detected; otherwise exits with 1.
    Auto-generated by Teams Background Packager for Intune.
#>
$ErrorActionPreference = 'Stop'

$AppName = "__APPNAME__"
$AppVersion = "__APPVERSION__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ExpectedImages = @(__IMAGEFILES__)

# Initialize Log Directory
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $fallbackDir = "$env:ProgramData\Log\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        try { New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {}
    }
}

$LogFile = Join-Path $LogDir "Detection.log"

function Write-DetectLog([string]$message) {
    $stamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $entry = "[$stamp] [DETECTION] $message"
    try { Add-Content -Path $LogFile -Value $entry -ErrorAction Stop } catch {}
}

$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

Write-DetectLog "Starting detection check for $AppName (v$AppVersion) under $($CurrentIdentity.Name)..."

function Test-AllImagesExist([string]$folder, [string]$folderLabel) {
    if (-not (Test-Path $folder)) {
        Write-DetectLog "Target folder not found ($folderLabel): $folder"
        return $false
    }
    $allPresent = $true
    foreach ($img in $ExpectedImages) {
        $imgPath = Join-Path $folder $img
        if (-not (Test-Path $imgPath)) {
            Write-DetectLog "Missing image ($folderLabel): $imgPath"
            $allPresent = $false
        } else {
            Write-DetectLog "Found image ($folderLabel): $img"
        }
    }
    return $allPresent
}

if ($IsSystemIdentity -or ($ConfiguredMode -eq "System")) {
    # Check if images exist in any human user profile under C:\Users (e.g. WDAGUtilityAccount in Sandbox) OR system marker
    Write-DetectLog "Running as SYSTEM / Sandbox identity. Checking user profiles under C:\Users..."
    
    $detectedInProfile = $false
    $userProfiles = try {
        Get-ChildItem -Path "C:\Users" -Directory -Force -ErrorAction Stop | Where-Object {
            $_.Name -notmatch '^(Public|All Users|Default User)$' -and
            -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
            ($_.Name -notmatch '^[\!\(\}]')
        }
    } catch { @() }

    foreach ($profile in $userProfiles) {
        if ($TargetNewTeams) {
            $userNewTeams = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            if (Test-AllImagesExist -folder $userNewTeams -folderLabel "Profile: $($profile.Name) - New Teams") {
                $detectedInProfile = $true
                break
            }
        }
        if ($TargetClassicTeams) {
            $userClassicTeams = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            if (Test-AllImagesExist -folder $userClassicTeams -folderLabel "Profile: $($profile.Name) - Classic Teams") {
                $detectedInProfile = $true
                break
            }
        }
    }

    $systemMarker = "$env:ProgramData\TeamsCustomBackgrounds\installed.json"
    $hasSystemMarker = (Test-Path $systemMarker)

    if ($detectedInProfile -or $hasSystemMarker) {
        Write-DetectLog "STATUS: INSTALLED (Verified via user profile / system marker under SYSTEM context)."
        Write-Output "Teams Backgrounds detected successfully ($($ExpectedImages.Count) images present)."
        exit 0
    } else {
        Write-DetectLog "STATUS: NOT INSTALLED (Images missing from user profiles and system marker missing)."
        Write-Warning "Teams Backgrounds not detected in user profiles."
        exit 1
    }
} else {
    # Current User Context
    $detected = $false

    if ($TargetNewTeams) {
        $newTeamsPath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
        if (Test-AllImagesExist -folder $newTeamsPath -folderLabel "New Teams") {
            $detected = $true
        }
    }

    if (-not $detected -and $TargetClassicTeams) {
        $classicTeamsPath = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
        if (Test-AllImagesExist -folder $classicTeamsPath -folderLabel "Classic Teams") {
            $detected = $true
        }
    }

    # Also check user manifest marker
    $markerFile = "$env:LOCALAPPDATA\TeamsCustomBackgrounds\installed.json"
    if (Test-Path $markerFile) {
        try {
            $marker = Get-Content $markerFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($marker.Version -eq "__APPVERSION__") {
                Write-DetectLog "Manifest marker verified: version $($marker.Version)"
                $detected = $true
            }
        } catch {}
    }

    if ($detected) {
        Write-DetectLog "STATUS: INSTALLED (All $($ExpectedImages.Count) images verified)."
        Write-Output "Teams Backgrounds detected successfully ($($ExpectedImages.Count) images present)."
        exit 0
    } else {
        Write-DetectLog "STATUS: NOT INSTALLED (Required background images missing)."
        Write-Warning "Teams Backgrounds not detected in user profile."
        exit 1
    }
}
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__APPVERSION__", $Version)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Function to generate the Uninstall script
function New-UninstallScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Uninstalls corporate Microsoft Teams background images.
    Auto-generated by Teams Background Packager for Intune.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$AppName = "__APPNAME__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ImageFiles = @(__IMAGEFILES__)

# Initialize Log Directory
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $fallbackDir = "$env:ProgramData\Log\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        try { New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {}
    }
}

$LogFile = Join-Path $LogDir "Uninstall.log"

function Write-UninstallLog([string]$message) {
    $stamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $entry = "[$stamp] [UNINSTALL] $message"
    Write-Output $entry
    try { Add-Content -Path $LogFile -Value $entry -ErrorAction Stop } catch {}
}

$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

Write-UninstallLog "========================================================"
Write-UninstallLog "Starting uninstallation of $AppName..."
Write-UninstallLog "Executing Identity: $($CurrentIdentity.Name)"
Write-UninstallLog "Host Computer:      $env:COMPUTERNAME"
Write-UninstallLog "Active Log File:    $LogFile"
Write-UninstallLog "========================================================"

function Remove-ImagesFromFolder([string]$folder, [string]$folderLabel) {
    if (-not (Test-Path $folder)) { return }
    foreach ($img in $ImageFiles) {
        $path = Join-Path $folder $img
        if (Test-Path $path) {
            try { Remove-Item -Path $path -Force -ErrorAction Stop } catch {}
            Write-UninstallLog "Removed ($folderLabel): $path"
        }
    }
}

if ($IsSystemIdentity -or ($ConfiguredMode -eq "System")) {
    Write-UninstallLog "Removing from all user profiles in C:\Users..."
    $userProfiles = try {
        Get-ChildItem -Path "C:\Users" -Directory -Force -ErrorAction Stop | Where-Object {
            $_.Name -notmatch '^(Public|All Users|Default User)$' -and
            -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
            ($_.Name -notmatch '^[\!\(\}]')
        }
    } catch { @() }

    foreach ($profile in $userProfiles) {
        if ($TargetNewTeams) {
            $p = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Remove-ImagesFromFolder -folder $p -folderLabel "Profile: $($profile.Name) - New Teams"
        }
        if ($TargetClassicTeams) {
            $p = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Remove-ImagesFromFolder -folder $p -folderLabel "Profile: $($profile.Name) - Classic Teams"
        }
        $userMarkerDir = Join-Path $profile.FullName "AppData\Local\TeamsCustomBackgrounds"
        if (Test-Path $userMarkerDir) {
            try { Remove-Item -Path $userMarkerDir -Recurse -Force -ErrorAction Stop } catch {}
        }
    }

    $markerDir = "$env:ProgramData\TeamsCustomBackgrounds"
    if (Test-Path $markerDir) {
        try { Remove-Item -Path $markerDir -Recurse -Force -ErrorAction Stop } catch {}
        Write-UninstallLog "Removed system deployment marker: $markerDir"
    }
} else {
    if ($TargetNewTeams) {
        $newTeamsUploads = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
        Remove-ImagesFromFolder -folder $newTeamsUploads -folderLabel "New Teams"
    }

    if ($TargetClassicTeams) {
        $classicTeamsUploads = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
        Remove-ImagesFromFolder -folder $classicTeamsUploads -folderLabel "Classic Teams"
    }

    $markerDir = "$env:LOCALAPPDATA\TeamsCustomBackgrounds"
    if (Test-Path $markerDir) {
        try { Remove-Item -Path $markerDir -Recurse -Force -ErrorAction Stop } catch {}
        Write-UninstallLog "Removed user deployment marker: $markerDir"
    }
}

Write-UninstallLog "Uninstallation completed successfully."
Write-Output "Teams Backgrounds removed successfully."
exit 0
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Main Packaging Action
$btnBuildPackage.Add_Click({
    # Validate inputs
    if ($Global:ImageItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please add at least one Teams background image before building the package.", "No Images Added", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $appName = $txtAppName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($appName)) {
        [System.Windows.MessageBox]::Show("Please enter an Application Name.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $targetNew = $chkNewTeams.IsChecked -eq $true
    $targetClassic = $chkClassicTeams.IsChecked -eq $true
    if (-not $targetNew -and -not $targetClassic) {
        [System.Windows.MessageBox]::Show("Please select at least one target Teams version (New Teams or Classic Teams).", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $outputDir = $txtOutputDir.Text.Trim()
    if (-not (Test-Path $outputDir)) {
        try {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Failed to create output directory: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
    }

    $deployMode = if ($rbSystemContext.IsChecked -eq $true) { "System" } elseif ($rbUserContext.IsChecked -eq $true) { "User" } else { "Auto" }
    $isUserContext = ($deployMode -eq "User")
    $version = $txtVersion.Text.Trim()
    if (-not $version) { $version = "1.0.0" }
    $publisher = $txtPublisher.Text.Trim()

    # Disable buttons during build
    $btnBuildPackage.IsEnabled = $false
    $btnAddImages.IsEnabled = $false
    $progressBar.Visibility = [System.Windows.Visibility]::Visible
    $progressBar.IsIndeterminate = $true
    $lblStatus.Text = "Checking Win32 Content Prep Tool (IntuneWinAppUtil.exe)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Check IntuneWinAppUtil
        $utilPath = Get-IntuneWinAppUtilPath
        if (-not $utilPath) {
            [System.Windows.MessageBox]::Show("IntuneWinAppUtil.exe is required to compile the .intunewin package. Build canceled.", "Missing Tool", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # Create Staging Directory
        $safeName = ($appName -replace '[^\w\-]', '_')
        $timestamp = (Get-Date).ToString("ddMMyy_HHmmss")
        $stagingRoot = Join-Path $outputDir "staging_${safeName}_$timestamp"
        $stagingSource = Join-Path $stagingRoot "Source"
        $stagingImages = Join-Path $stagingSource "Images"

        New-Item -Path $stagingImages -ItemType Directory -Force | Out-Null

        $lblStatus.Text = "Copying background images to package source..."
        [System.Windows.Forms.Application]::DoEvents()

        $imageNames = @()
        foreach ($item in $Global:ImageItems) {
            $destFile = Join-Path $stagingImages $item.FileName
            Copy-Item -Path $item.FullPath -Destination $destFile -Force
            $imageNames += $item.FileName
        }

        $lblStatus.Text = "Generating installation, uninstallation, and detection scripts..."
        [System.Windows.Forms.Application]::DoEvents()

        # Generate scripts
        $logDir = $txtLogDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($logDir)) {
            $logDir = "C:\Logs\$safeName"
            $txtLogDir.Text = $logDir
        }

        # Generate scripts with full enterprise logging
        $installScript   = New-InstallScriptContent   -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -Version $version -AppName $appName -LogDir $logDir
        $detectScript    = New-DetectScriptContent    -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -Version $version -AppName $appName -LogDir $logDir
        $uninstallScript = New-UninstallScriptContent -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -AppName $appName -LogDir $logDir

        # Save into Staging Source
        $installPath = Join-Path $stagingSource "Install-TeamsBackgrounds.ps1"
        $detectPath = Join-Path $stagingSource "Detect-TeamsBackgrounds.ps1"
        $uninstallPath = Join-Path $stagingSource "Uninstall-TeamsBackgrounds.ps1"

        Set-Content -Path $installPath -Value $installScript -Encoding UTF8 -Force
        Set-Content -Path $detectPath -Value $detectScript -Encoding UTF8 -Force
        Set-Content -Path $uninstallPath -Value $uninstallScript -Encoding UTF8 -Force

        # Also save detection script and install instructions directly to Output directory for easy access
        $outputDetectScript = Join-Path $outputDir "Detect-TeamsBackgrounds.ps1"
        Set-Content -Path $outputDetectScript -Value $detectScript -Encoding UTF8 -Force
        $Global:LastBuiltDetectionScriptPath = $outputDetectScript

        $lblStatus.Text = "Compiling Win32 application package (.intunewin)..."
        [System.Windows.Forms.Application]::DoEvents()

        # Run IntuneWinAppUtil.exe
        $processArgs = @(
            "-c", "`"$stagingSource`"",
            "-s", "`"Install-TeamsBackgrounds.ps1`"",
            "-o", "`"$outputDir`"",
            "-q"
        ) -join " "

        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $utilPath
        $pinfo.Arguments = $processArgs
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($pinfo)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        # Expected output package name is Install-TeamsBackgrounds.intunewin
        $defaultIntuneWin = Join-Path $outputDir "Install-TeamsBackgrounds.intunewin"
        $customIntuneWin = Join-Path $outputDir "${safeName}.intunewin"

        if (Test-Path $defaultIntuneWin) {
            # Rename to application-specific name
            if (Test-Path $customIntuneWin) { Remove-Item -Path $customIntuneWin -Force }
            Move-Item -Path $defaultIntuneWin -Destination $customIntuneWin -Force
            $Global:LastBuiltPackagePath = $customIntuneWin
        } elseif (Test-Path $customIntuneWin) {
            $Global:LastBuiltPackagePath = $customIntuneWin
        } else {
            # Search output directory for any newly created intunewin
            $recent = Get-ChildItem -Path $outputDir -Filter "*.intunewin" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($recent) {
                $Global:LastBuiltPackagePath = $recent.FullName
            } else {
                throw "Packaging failed. IntuneWinAppUtil output: $stdout $stderr"
            }
        }

        # Clean up staging directory
        try {
            Remove-Item -Path $stagingRoot -Recurse -Force -ErrorAction Stop
        } catch {}

        # Save Intune Setup Instructions in output folder aligned with the Intune upload wizard
        $behaviorStr = if ($isUserContext) { 'User' } else { 'System' }
        $packageFileName = [System.IO.Path]::GetFileName($Global:LastBuiltPackagePath)
        $generatedTimestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
        $createdDateStr = Get-Date -Format "dd/MM/yy"
        $clientsSummary = @(
            if ($targetNew) { "New Teams (v2)" }
            if ($targetClassic) { "Classic Teams (v1)" }
        ) -join ", "
        $instructionsPath = Join-Path $outputDir "Intune_Configuration_Guide.txt"
        $instructions = @"
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : $generatedTimestamp
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
$packageFileName


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
$appName

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
$($txtDescription.Text)
All runtime logs are saved to $logDir\

PACKAGE DETAILS:
- Application Name : $appName
- Target Clients   : $clientsSummary
- Image Count      : $($imageNames.Count) image(s)
- Author           : JP
- App Version      : $version
- Created          : $createdDateStr

WORKSTATION LOG LOCATION:
$logDir\

LOG FILES IN THIS FOLDER:
- Install Log   : $logDir\Install.log
- Detection Log : $logDir\Detection.log
- Uninstall Log : $logDir\Uninstall.log
======================================================================

Publisher:
$publisher

App Version:
$version


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install-TeamsBackgrounds.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall-TeamsBackgrounds.ps1

Install behavior:
$behaviorStr

Device restart behavior:
Determine behavior based on return codes


STEP 4: REQUIREMENTS
------------------------------------------------------------------------
Operating system architecture:
64-bit

Minimum operating system:
Windows 10 1909


STEP 5: DETECTION RULES
------------------------------------------------------------------------
Rules format:
Use a custom detection script

Script file:
Detect-TeamsBackgrounds.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No


========================================================================
FILES IN THIS OUTPUT FOLDER
========================================================================
- $packageFileName              : Upload to Intune Win32 App Package
- Detect-TeamsBackgrounds.ps1      : Upload to Intune Custom Detection Script
- Intune_Configuration_Guide.txt   : This deployment reference guide
========================================================================
"@
        Set-Content -Path $instructionsPath -Value $instructions -Encoding UTF8 -Force

        # Update Intune Guidance tab UI
        $lblPackagePath.Text = $Global:LastBuiltPackagePath
        $txtInstallCommand.Text = 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install-TeamsBackgrounds.ps1'
        $txtUninstallCommand.Text = 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall-TeamsBackgrounds.ps1'
        $txtInstallBehavior.Text = if ($isUserContext) { "User" } else { "System" }
        $txtDetectionScriptPreview.Text = $detectScript
        $lblLogDirPath.Text = $logDir
        $lblInstallLogPath.Text = Join-Path $logDir "Install.log"
        $lblDetectLogPath.Text = Join-Path $logDir "Detection.log"
        $lblUninstallLogPath.Text = Join-Path $logDir "Uninstall.log"

        # Write Packager Tool Log
        try {
            $packagerLogDir = "C:\Logs\TeamsBackgroundPackager"
            if (-not (Test-Path $packagerLogDir)) { New-Item -Path $packagerLogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
            $packagerLogFile = Join-Path $packagerLogDir "Packager.log"
            $pStamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
            Add-Content -Path $packagerLogFile -Value "[$pStamp] [BUILD_SUCCESS] Package: '$appName' (v$version) -> '$Global:LastBuiltPackagePath' | Images: $($imageNames.Count) | Endpoint LogDir: '$logDir'" -ErrorAction Stop
        } catch {}

        $lblStatus.Text = "Success! Package created: $([System.IO.Path]::GetFileName($Global:LastBuiltPackagePath))"
        
        # Switch to Intune Guidance Tab
        $MainTabs.SelectedItem = $tabIntuneGuidance

        [System.Windows.MessageBox]::Show("Win32 package successfully created!`n`nFile: $([System.IO.Path]::GetFileName($Global:LastBuiltPackagePath))`n`nReview the 'Intune Deployment Details' tab for the exact command lines and detection rules.", "Package Built Successfully", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

    } catch {
        $lblStatus.Text = "Error building package."
        [System.Windows.MessageBox]::Show("An error occurred during build:`n$_", "Build Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        $btnBuildPackage.IsEnabled = $true
        $btnAddImages.IsEnabled = $true
        $progressBar.Visibility = [System.Windows.Visibility]::Collapsed
        $progressBar.IsIndeterminate = $false
    }
})

# Copy buttons actions
$btnCopyInstallCmd.Add_Click({
    if ($txtInstallCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtInstallCommand.Text)
        $lblStatus.Text = "Install command copied to clipboard!"
    }
})

$btnCopyUninstallCmd.Add_Click({
    if ($txtUninstallCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtUninstallCommand.Text)
        $lblStatus.Text = "Uninstall command copied to clipboard!"
    }
})

$btnCopyDetectionScript.Add_Click({
    if ($txtDetectionScriptPreview.Text) {
        [System.Windows.Clipboard]::SetText($txtDetectionScriptPreview.Text)
        $lblStatus.Text = "Detection script copied to clipboard!"
    }
})

$btnCopyPackagePath.Add_Click({
    if ($Global:LastBuiltPackagePath) {
        [System.Windows.Clipboard]::SetText($Global:LastBuiltPackagePath)
        $lblStatus.Text = "Package path copied to clipboard!"
    }
})

$btnOpenOutputDir.Add_Click({
    $dir = $txtOutputDir.Text
    if (Test-Path $dir) {
        Start-Process explorer.exe -ArgumentList "`"$dir`""
    }
})

$btnOpenDetectionFile.Add_Click({
    if ($Global:LastBuiltDetectionScriptPath -and (Test-Path $Global:LastBuiltDetectionScriptPath)) {
        Start-Process notepad.exe -ArgumentList "`"$Global:LastBuiltDetectionScriptPath`""
    }
})

$btnOpenLogDir.Add_Click({
    $dir = $txtLogDir.Text.Trim()
    if (-not (Test-Path $dir)) {
        try { New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {}
    }
    if (Test-Path $dir) {
        Start-Process explorer.exe -ArgumentList "`"$dir`""
    }
})

# Auto-update Log Directory when Application Name changes
$txtAppName.Add_TextChanged({
    $safe = ($txtAppName.Text.Trim() -replace '[^\w\-]', '_')
    if (-not $safe) { $safe = "TeamsBackgrounds" }
    $txtLogDir.Text = "C:\Logs\$safe"
})

# Launch GUI
$Window.ShowDialog() | Out-Null
