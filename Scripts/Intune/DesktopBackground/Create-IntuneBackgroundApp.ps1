<#
.SYNOPSIS
    WPF GUI utility for building, staging, and packaging corporate Desktop Background and Lock Screen deployments for Intune.
.DESCRIPTION
    Launches a modern WPF interface with a painted Windows 11 titlebar, automatic system Light/Dark theme support,
    and a live 16:9 Desktop and Lock Screen preview. Supports local image files or dynamic remote URLs, generates
    staging folders, installation scripts (Install.ps1, ApplyBackground.ps1, DownloadBackgrounds.ps1, Uninstall.ps1),
    custom detection scripts, and packs everything into a ready-to-upload .intunewin bundle with a 5-step guide.
.PARAMETER DebugMode
    Launches the packaging process with verbose execution logs.
.EXAMPLE
    .\Create-IntuneBackgroundApp.ps1
.NOTES
    Script Name   : Create-IntuneBackgroundApp.ps1
    Author        : JP
    Version       : 2.0
    Created       : 17/07/26
    Last Modified : 26/11/26
    Log Path      : C:\Logs\DesktopBackground\
#>
[CmdletBinding()]
param(
    [switch]$DebugMode = $false
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Enforce STA Threading for WPF (Supports Local & In-Memory irm | iex)
# ---------------------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $relaunchPath = if ($PSScriptRoot -and $MyInvocation.MyCommand.Name) {
        Join-Path -Path $PSScriptRoot -ChildPath $MyInvocation.MyCommand.Name
    } elseif ($PSCommandPath) {
        $PSCommandPath
    } else {
        $null
    }

    if (-not $relaunchPath -or -not (Test-Path -Path $relaunchPath -PathType Leaf)) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
        $cacheRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'IntuneDesktopBackground'
        if (-not (Test-Path -Path $cacheRoot)) {
            New-Item -Path $cacheRoot -ItemType Directory -Force | Out-Null
        }
        $relaunchPath = Join-Path -Path $cacheRoot -ChildPath 'Create-IntuneBackgroundApp.ps1'
        Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Intune/DesktopBackground/Create-IntuneBackgroundApp.ps1' -OutFile $relaunchPath -UseBasicParsing
    }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-STA')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-NoProfile')
    $argList.Add('-File')
    $argList.Add("`"$relaunchPath`"")
    if ($DebugMode) {
        $argList.Add('-DebugMode')
    }
    Start-Process -FilePath 'powershell.exe' -ArgumentList ($argList -join ' ')
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Xaml

# ---------------------------------------------------------------------------
# 2. Painted Windows 11 Chrome Helper (Frame / DwmSetWindowAttribute Pattern)
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# 3. Theme Detection, Figtree Font Loader & Standard 70% x 80% Window Sizing
# ---------------------------------------------------------------------------
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
            BrushScrollThumb   = '#26354D'
            BrushSwitchOff     = '#26354D'
            BrushSwitchKnobOn  = '#042F2E'
            BrushSwitchKnobOff = '#F8FAFC'
            BrushOk            = '#34D399'
            BrushErr           = '#F87171'
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
            BrushScrollThumb   = '#CBD5E1'
            BrushSwitchOff     = '#CBD5E1'
            BrushSwitchKnobOn  = '#FFFFFF'
            BrushSwitchKnobOff = '#FFFFFF'
            BrushOk            = '#059669'
            BrushErr           = '#DC2626'
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
        $fontUri = ([System.Uri](($resolvedFontDir.TrimEnd('\') + '\'))).AbsoluteUri
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
    $Window.Width      = $targetWidth
    $Window.Height     = $targetHeight
    $Window.MinWidth   = [Math]::Min(860, [int]$targetWidth)
    $Window.MinHeight  = [Math]::Min(640, [int]$targetHeight)
    $Window.Left       = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $Window.Top        = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
}

# ---------------------------------------------------------------------------
# 4. Logging & IntuneWinAppUtil Locator
# ---------------------------------------------------------------------------
$Script:LogTextBox = $null
$Script:FooterStatusText = $null

function Invoke-UiRefresh {
    [CmdletBinding()]
    param()
    if ([System.Windows.Application]::Current -or $null -ne $Script:LogTextBox) {
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [Action[System.Windows.Threading.DispatcherFrame]]{ param($f) $f.Continue = $false },
            $frame
        ) | Out-Null
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    }
}

function Write-BuildLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $logDir = 'C:\Logs\DesktopBackground'
    if (-not (Test-Path -Path $logDir)) {
        try {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            $logDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Logs\DesktopBackground'
            if (-not (Test-Path -Path $logDir)) {
                try { New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch {}
            }
        }
    }

    $timestamp = Get-Date -Format 'dd/MM/yy HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    try {
        Add-Content -Path (Join-Path -Path $logDir -ChildPath 'Packager.log') -Value $entry -Encoding UTF8 -ErrorAction Stop
    } catch {}

    if ($null -ne $Script:LogTextBox) {
        $Script:LogTextBox.AppendText("$entry`r`n")
        $Script:LogTextBox.ScrollToEnd()
        Invoke-UiRefresh
    } else {
        Write-Verbose $entry
    }
}

function Get-IntuneWinAppUtilPath {
    [CmdletBinding()]
    param(
        [System.Windows.Controls.TextBlock]$StatusTextBlock = $null
    )

    $candidatePaths = @()
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidatePaths += (Join-Path $PSScriptRoot 'IntuneWinAppUtil.exe')
        $candidatePaths += (Join-Path $PSScriptRoot 'tools\IntuneWinAppUtil.exe')
    }

    $localAppDataDir = Join-Path $env:LOCALAPPDATA 'Microsoft\IntuneWinAppUtil'
    $customAppDataDir = Join-Path $env:LOCALAPPDATA 'IntuneDesktopBackground'
    $printerAppDataDir = Join-Path $env:LOCALAPPDATA 'IntunePrinterPackager'
    $candidatePaths += (Join-Path $localAppDataDir 'IntuneWinAppUtil.exe')
    $candidatePaths += (Join-Path $customAppDataDir 'IntuneWinAppUtil.exe')
    $candidatePaths += (Join-Path $printerAppDataDir 'IntuneWinAppUtil.exe')
    $candidatePaths += (Join-Path $env:TEMP 'IntuneWinAppUtil.exe')

    $userDownloads = Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) 'Downloads'
    $candidatePaths += (Join-Path $userDownloads 'IntuneWinAppUtil.exe')

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $intuneRoot = Split-Path -Parent $PSScriptRoot
        if (-not [string]::IsNullOrWhiteSpace($intuneRoot)) {
            $candidatePaths += (Join-Path $intuneRoot 'AddPrinter\IntuneWinAppUtil.exe')
            $candidatePaths += (Join-Path $intuneRoot 'DesktopOverlay\IntuneWinAppUtil.exe')
            $candidatePaths += (Join-Path $intuneRoot 'TeamsBackground\tools\IntuneWinAppUtil.exe')
        }
    }

    $cmd = try { Get-Command 'IntuneWinAppUtil.exe' -ErrorAction Stop } catch { $null }
    if ($cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source) -and (Test-Path -Path $cmd.Source -PathType Leaf)) {
        return $cmd.Source
    }

    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -Path $cand -PathType Leaf)) {
            try {
                if ((Get-Item $cand).Length -gt 10000) {
                    return $cand
                }
            } catch {}
        }
    }

    $downloadTargetDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
        $PSScriptRoot
    } else {
        $customAppDataDir
    }
    if (-not (Test-Path -Path $downloadTargetDir)) {
        try { New-Item -ItemType Directory -Path $downloadTargetDir -Force -ErrorAction Stop | Out-Null } catch {}
    }
    $targetExePath = Join-Path $downloadTargetDir 'IntuneWinAppUtil.exe'

    $msg = "Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) was not found.`n`nWould you like to automatically download it from Microsoft's official GitHub repository?"
    $confirm = [System.Windows.MessageBox]::Show($msg, 'Packer Tool Missing', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        if ($null -ne $StatusTextBlock) {
            $StatusTextBlock.Text = 'Downloading IntuneWinAppUtil.exe...'
            Invoke-UiRefresh
        }
        Write-BuildLog -Message 'Downloading IntuneWinAppUtil.exe from Microsoft GitHub...' -Level INFO
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
            $downloadUrl = 'https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe'
            $prevProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $targetExePath -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            } finally {
                $ProgressPreference = $prevProgress
            }
            $sig = Get-AuthenticodeSignature -FilePath $targetExePath -ErrorAction Stop
            if ($sig.Status -ne 'Valid') {
                if (Test-Path -LiteralPath $targetExePath) { try { Remove-Item $targetExePath -Force -ErrorAction Stop } catch {} }
                throw "Authenticode signature check failed for IntuneWinAppUtil.exe ($($sig.Status))"
            }
            if (Test-Path -Path $targetExePath -PathType Leaf) {
                Write-BuildLog -Message "Downloaded and verified IntuneWinAppUtil.exe at $targetExePath" -Level SUCCESS
                return $targetExePath
            }
            throw 'Downloaded file was not found on disk.'
        } catch {
            Write-BuildLog -Message "Failed to download IntuneWinAppUtil.exe: $($_.Exception.Message)" -Level ERROR
            [System.Windows.MessageBox]::Show("Failed to download IntuneWinAppUtil.exe:`n`n$_`n`nPlease download it manually and place it in: $downloadTargetDir", 'Download Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return $null
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# 5. Modern WPF XAML UI (Dynamic System Dark/Light + Uniform Control Grid)
# ---------------------------------------------------------------------------
$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Desktop Background Packager v2.0"
        WindowStartupLocation="Manual"
        ResizeMode="CanResize"
        Background="#091122"
        Foreground="{DynamicResource BrushTextPrimary}"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True">
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
        <SolidColorBrush x:Key="BrushScrollThumb" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushSwitchOff" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushSwitchKnobOn" Color="#042F2E"/>
        <SolidColorBrush x:Key="BrushSwitchKnobOff" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BrushOk" Color="#34D399"/>
        <SolidColorBrush x:Key="BrushErr" Color="#F87171"/>

        <!-- Slim ScrollBar -->
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

        <!-- Clean Input TextBox (Uniform 32px) -->
        <Style x:Key="FieldTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Background" Value="{DynamicResource BrushInputBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="CaretBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary CTA Button -->
        <Style x:Key="PrimaryCtaStyle" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BrushAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button -->
        <Style x:Key="SecondaryBtnStyle" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Pill Toggle Switch (for CheckBox sections) -->
        <Style x:Key="ToggleSwitchStyle" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Border x:Name="Track" Width="34" Height="19" CornerRadius="9.5" Background="{DynamicResource BrushSwitchOff}" VerticalAlignment="Center">
                                <Ellipse x:Name="Knob" Width="13" Height="13" Fill="{DynamicResource BrushSwitchKnobOff}" HorizontalAlignment="Left" Margin="3,0,0,0"/>
                            </Border>
                            <ContentPresenter Margin="9,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Track" Property="Background" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="Knob" Property="HorizontalAlignment" Value="Right"/>
                                <Setter TargetName="Knob" Property="Margin" Value="0,0,3,0"/>
                                <Setter TargetName="Knob" Property="Fill" Value="{DynamicResource BrushSwitchKnobOn}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Clean Dark/Light CheckBox -->
        <Style x:Key="ModernCheckBoxStyle" TargetType="CheckBox">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Border x:Name="Box"
                                    Width="16" Height="16"
                                    CornerRadius="4"
                                    Background="{DynamicResource BrushInputBg}"
                                    BorderBrush="{DynamicResource BrushCardBorder}"
                                    BorderThickness="1"
                                    VerticalAlignment="Center">
                                <Path x:Name="Tick"
                                      Data="M 3,8 L 6.5,11.5 L 13,4.5"
                                      Stroke="{DynamicResource BrushAccentFg}"
                                      StrokeThickness="2"
                                      StrokeStartLineCap="Round"
                                      StrokeEndLineCap="Round"
                                      StrokeLineJoin="Round"
                                      Visibility="Collapsed"/>
                            </Border>
                            <ContentPresenter Margin="8,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box" Property="Background" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="Tick" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Box" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Equal-Width Segmented RadioButton -->
        <Style x:Key="SegmentedRadioStyle" TargetType="RadioButton">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextMuted}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="FontSize" Value="11.5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="Pill" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="1" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Pill" Property="Background" Value="{DynamicResource BrushAccentTint}"/>
                                <Setter TargetName="Pill" Property="BorderBrush" Value="{DynamicResource BrushAccentBorder}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Section Card Container -->
        <Style x:Key="SectionCardStyle" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource BrushCardBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="7"/>
            <Setter Property="Padding" Value="13,11"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Band 1: Header Bar (Tool-Only Name & Version + Auto System Theme Toggle) -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="20,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,16,0">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Intune Desktop Background Packager" FontSize="16" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextPrimary}"/>
                        <Border Margin="10,0,0,0" Background="{DynamicResource BrushAccentTint}" BorderBrush="{DynamicResource BrushAccentBorder}" BorderThickness="1" CornerRadius="4" Padding="7,2" VerticalAlignment="Center">
                            <TextBlock x:Name="TxtVersionPill" Text="v2.0" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                        </Border>
                    </StackPanel>
                    <TextBlock Margin="0,3,0,0" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" TextTrimming="CharacterEllipsis"
                               Text="Package corporate Desktop Wallpaper and Lock Screen branding into a Win32 .intunewin bundle."/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnThemeToggle" Style="{StaticResource SecondaryBtnStyle}" Margin="0,0,12,0" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock x:Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <TextBlock Text="Ver" FontSize="12" FontWeight="Medium" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center" Margin="0,0,7,0"/>
                    <TextBox x:Name="TxtVersion" Style="{StaticResource FieldTextBoxStyle}" Width="58" Padding="4,0" HorizontalContentAlignment="Center" FontFamily="Cascadia Mono, Consolas" FontWeight="SemiBold" Text="2.0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Band 2: Flush 2-Column Workspace -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.12*"/>
                <ColumnDefinition Width="0.88*"/>
            </Grid.ColumnDefinitions>

            <!-- Left Column: Inputs (Desktop Background + Lock Screen + Output Directory) -->
            <Border Grid.Column="0" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,1,0" Padding="18,14">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Card 1: Desktop Background -->
                    <Border x:Name="CardDesktop" Grid.Row="0" Style="{StaticResource SectionCardStyle}" AllowDrop="True" Margin="0,0,0,12">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="196"/>
                                </Grid.ColumnDefinitions>
                                <CheckBox x:Name="ChkSetDesktop" Grid.Column="0" Style="{StaticResource ToggleSwitchStyle}" IsChecked="True" Content="Desktop Background"/>
                                <Border Grid.Column="1" Height="30" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="2">
                                    <UniformGrid Columns="2">
                                        <RadioButton x:Name="RadLocal" Style="{StaticResource SegmentedRadioStyle}" GroupName="DeskSource" IsChecked="True" Content="Local File"/>
                                        <RadioButton x:Name="RadUrl" Style="{StaticResource SegmentedRadioStyle}" GroupName="DeskSource" Content="Remote URL"/>
                                    </UniformGrid>
                                </Border>
                            </Grid>

                            <Border x:Name="PnlDeskBody" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Margin="0,10,0,0" Padding="0,10,0,0">
                                <StackPanel>
                                    <!-- Local File Row -->
                                    <Grid x:Name="PnlDeskLocal" Margin="0,0,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="92"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox x:Name="TxtLocalFile" Grid.Column="0" Style="{StaticResource FieldTextBoxStyle}" FontFamily="Cascadia Mono, Consolas" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnBrowseLocal" Grid.Column="1" Style="{StaticResource SecondaryBtnStyle}" Content="Browse…"/>
                                    </Grid>

                                    <!-- Remote URL Row -->
                                    <Grid x:Name="PnlDeskUrl" Visibility="Collapsed" Margin="0,0,0,10">
                                        <TextBox x:Name="TxtUrl" Style="{StaticResource FieldTextBoxStyle}" FontFamily="Cascadia Mono, Consolas" ToolTip="HTTPS URL of wallpaper downloaded and cached on endpoints"/>
                                    </Grid>

                                    <!-- Full-Width 6-Column Wallpaper Style Selector -->
                                    <TextBlock Text="Wallpaper Style" FontSize="11.5" FontWeight="Medium" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,5"/>
                                    <Border Height="32" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="2" Margin="0,0,0,10">
                                        <UniformGrid Columns="6">
                                            <RadioButton x:Name="RadStyleFill" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" IsChecked="True" Content="Fill"/>
                                            <RadioButton x:Name="RadStyleFit" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" Content="Fit"/>
                                            <RadioButton x:Name="RadStyleStretch" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" Content="Stretch"/>
                                            <RadioButton x:Name="RadStyleTile" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" Content="Tile"/>
                                            <RadioButton x:Name="RadStyleCenter" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" Content="Center"/>
                                            <RadioButton x:Name="RadStyleSpan" Style="{StaticResource SegmentedRadioStyle}" GroupName="WallpaperStyle" Content="Span"/>
                                        </UniformGrid>
                                    </Border>

                                    <CheckBox x:Name="ChkLockBackground" Style="{StaticResource ModernCheckBoxStyle}" Content="Prevent users from changing desktop wallpaper"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>

                    <!-- Card 2: Lock Screen -->
                    <Border x:Name="CardLockScreen" Grid.Row="1" Style="{StaticResource SectionCardStyle}" AllowDrop="True" Margin="0,0,0,12">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <CheckBox x:Name="ChkSetLockScreen" Grid.Column="0" Style="{StaticResource ToggleSwitchStyle}" IsChecked="True" Content="Lock Screen"/>
                                <CheckBox x:Name="ChkSameAsDesk" Grid.Column="1" Style="{StaticResource ModernCheckBoxStyle}" IsChecked="True" Content="Same as Desktop" FontWeight="Medium"/>
                            </Grid>

                            <Border x:Name="PnlLockCustom" Visibility="Collapsed" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Margin="0,10,0,0" Padding="0,10,0,0">
                                <StackPanel>
                                    <Grid Margin="0,0,0,8">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="196"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Grid.Column="0" Text="Lock Screen Image" FontSize="11.5" FontWeight="Medium" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center"/>
                                        <Border Grid.Column="1" Height="30" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="2">
                                            <UniformGrid Columns="2">
                                                <RadioButton x:Name="RadLockLocal" Style="{StaticResource SegmentedRadioStyle}" GroupName="LockSource" IsChecked="True" Content="Local File"/>
                                                <RadioButton x:Name="RadLockUrl" Style="{StaticResource SegmentedRadioStyle}" GroupName="LockSource" Content="Remote URL"/>
                                            </UniformGrid>
                                        </Border>
                                    </Grid>

                                    <Grid x:Name="PnlLockLocal">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="92"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox x:Name="TxtLockLocalFile" Grid.Column="0" Style="{StaticResource FieldTextBoxStyle}" FontFamily="Cascadia Mono, Consolas" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnBrowseLockLocal" Grid.Column="1" Style="{StaticResource SecondaryBtnStyle}" Content="Browse…"/>
                                    </Grid>

                                    <Grid x:Name="PnlLockUrl" Visibility="Collapsed">
                                        <TextBox x:Name="TxtLockUrl" Style="{StaticResource FieldTextBoxStyle}" FontFamily="Cascadia Mono, Consolas"/>
                                    </Grid>
                                </StackPanel>
                            </Border>

                            <CheckBox x:Name="ChkLockLockScreen" Style="{StaticResource ModernCheckBoxStyle}" Margin="0,10,0,0" Content="Prevent users from changing lock screen"/>
                        </StackPanel>
                    </Border>

                    <!-- Card 3: Output Directory (Anchored cleanly at bottom of Left Column) -->
                    <Border Grid.Row="3" Style="{StaticResource SectionCardStyle}">
                        <StackPanel>
                            <Grid Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="OUTPUT DIRECTORY" FontSize="11.5" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                                <TextBlock Grid.Column="1" Text="Exports .intunewin, detection &amp; guide" FontSize="11" Foreground="{DynamicResource BrushTextMuted}"/>
                            </Grid>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="92"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="TxtOutputDir" Grid.Column="0" Style="{StaticResource FieldTextBoxStyle}" FontFamily="Cascadia Mono, Consolas" Margin="0,0,8,0"/>
                                <Button x:Name="BtnBrowseOutput" Grid.Column="1" Style="{StaticResource SecondaryBtnStyle}" Content="Browse…"/>
                            </Grid>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>

            <!-- Right Column: Live 16:9 Screen Preview + Build Log -->
            <Grid Grid.Column="1" Margin="18,14">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Card 1: Live Screen Preview (Desktop & Lock Screen) -->
                <Border x:Name="CardPreview" Grid.Row="0" Style="{StaticResource SectionCardStyle}" AllowDrop="True" Margin="0,0,0,12">
                    <StackPanel>
                        <Grid Margin="0,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="196"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="SCREEN PREVIEW" FontSize="12" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}" VerticalAlignment="Center"/>
                                <Border Margin="8,0,0,0" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="4" Padding="6,2" VerticalAlignment="Center">
                                    <TextBlock x:Name="TxtPreviewBadge" Text="Fill · 16:9" FontFamily="Cascadia Mono, Consolas" FontSize="10.5" Foreground="{DynamicResource BrushTextMuted}"/>
                                </Border>
                            </StackPanel>
                            <Border Grid.Column="1" Height="30" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="2">
                                <UniformGrid Columns="2">
                                    <RadioButton x:Name="RadPrevDesktop" Style="{StaticResource SegmentedRadioStyle}" GroupName="PreviewTarget" IsChecked="True" Content="Desktop"/>
                                    <RadioButton x:Name="RadPrevLock" Style="{StaticResource SegmentedRadioStyle}" GroupName="PreviewTarget" Content="Lock Screen"/>
                                </UniformGrid>
                            </Border>
                        </Grid>

                        <!-- 16:9 Monitor Bezel -->
                        <Border Background="#04070F" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" ClipToBounds="True">
                            <Viewbox Stretch="Uniform">
                                <Grid Width="480" Height="270" Background="#04070F">
                                    <Image x:Name="ImgPreview" Stretch="UniformToFill" RenderOptions.BitmapScalingMode="HighQuality"/>
                                    <Rectangle x:Name="RectPreviewTile" Visibility="Collapsed"/>
                                    <StackPanel x:Name="PnlPreviewEmpty" HorizontalAlignment="Center" VerticalAlignment="Center">
                                        <TextBlock Text="Select or drop a local wallpaper image to preview" FontSize="12.5" Foreground="#94A3B8" HorizontalAlignment="Center"/>
                                        <TextBlock Text="Supports PNG, JPG, JPEG, BMP" FontSize="11" Foreground="#64748B" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                                    </StackPanel>

                                    <!-- Subtle Windows 11 Desktop Taskbar Overlay (shown once an image is loaded) -->
                                    <Grid x:Name="PnlOverlayDesktop" Visibility="Collapsed" IsHitTestVisible="False">
                                        <Border Height="24" VerticalAlignment="Bottom" Background="#A6091122" BorderBrush="#2AFFFFFF" BorderThickness="0,1,0,0">
                                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                                                <Border Width="11" Height="11" CornerRadius="2" Background="#38BDF8" Margin="0,0,8,0"/>
                                                <Border Width="54" Height="9" CornerRadius="4.5" Background="#2AFFFFFF" Margin="0,0,8,0"/>
                                                <Border Width="10" Height="10" CornerRadius="2" Background="#2DD4BF" Margin="0,0,6,0"/>
                                                <Border Width="10" Height="10" CornerRadius="2" Background="#FBBF24"/>
                                            </StackPanel>
                                        </Border>
                                    </Grid>

                                    <!-- Windows 11 Lock Screen Clock Overlay -->
                                    <StackPanel x:Name="PnlOverlayLock" Visibility="Collapsed" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,30,0,0" IsHitTestVisible="False">
                                        <TextBlock x:Name="TxtLockClock" Text="21:24" FontSize="36" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center">
                                            <TextBlock.Effect>
                                                <DropShadowEffect BlurRadius="10" ShadowDepth="2" Opacity="0.65"/>
                                            </TextBlock.Effect>
                                        </TextBlock>
                                        <TextBlock x:Name="TxtLockDate" Text="Thursday, 26 November" FontSize="12.5" FontWeight="SemiBold" Foreground="#F8FAFC" HorizontalAlignment="Center" Margin="0,2,0,0">
                                            <TextBlock.Effect>
                                                <DropShadowEffect BlurRadius="8" ShadowDepth="1" Opacity="0.65"/>
                                            </TextBlock.Effect>
                                        </TextBlock>
                                    </StackPanel>
                                </Grid>
                            </Viewbox>
                        </Border>
                    </StackPanel>
                </Border>

                <!-- Card 2: Always-Visible Build Log Console -->
                <Border Grid.Row="1" Style="{StaticResource SectionCardStyle}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <Grid Grid.Row="0" Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="BUILD LOG" FontSize="11.5" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}" VerticalAlignment="Center"/>
                            <TextBlock Grid.Column="1" Text="C:\Logs\DesktopBackground\" FontFamily="Cascadia Mono, Consolas" FontSize="10.5" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center"/>
                        </Grid>

                        <Border Grid.Row="1" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="8">
                            <TextBox x:Name="TxtBuildLog"
                                     Background="Transparent"
                                     Foreground="{DynamicResource BrushTextSecondary}"
                                     BorderThickness="0"
                                     FontFamily="Cascadia Mono, Consolas"
                                     FontSize="11.5"
                                     IsReadOnly="True"
                                     TextWrapping="Wrap"
                                     AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto"
                                     Text=""/>
                        </Border>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- Band 3: Single-Row Footer (Left Status Hint | Right Action Cluster) -->
        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="20,11">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="TxtFooterStatus" Grid.Column="0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource BrushTextMuted}" TextTrimming="CharacterEllipsis" Margin="0,0,16,0"
                           Text="Ready. Configure Desktop Background or Lock Screen options and click Generate Intune Package."/>

                <UniformGrid Grid.Column="1" Columns="2" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnOpenOutput" Style="{StaticResource SecondaryBtnStyle}" Content="Open Output Folder" Margin="0,0,5,0"/>
                    <Button x:Name="BtnGenerate" Style="{StaticResource PrimaryCtaStyle}" Content="Generate Intune Package" Margin="5,0,0,0"/>
                </UniformGrid>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xamlText)
$window = [Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $window

# Bind Named Controls
$txtVersionPill     = $window.FindName('TxtVersionPill')
$txtVersion         = $window.FindName('TxtVersion')
$btnThemeToggle     = $window.FindName('BtnThemeToggle')
$txtThemeLabel      = $window.FindName('TxtThemeLabel')

$cardDesktop        = $window.FindName('CardDesktop')
$chkSetDesktop      = $window.FindName('ChkSetDesktop')
$radLocal           = $window.FindName('RadLocal')
$radUrl             = $window.FindName('RadUrl')
$pnlDeskBody        = $window.FindName('PnlDeskBody')
$pnlDeskLocal       = $window.FindName('PnlDeskLocal')
$pnlDeskUrl         = $window.FindName('PnlDeskUrl')
$txtLocalFile       = $window.FindName('TxtLocalFile')
$btnBrowseLocal     = $window.FindName('BtnBrowseLocal')
$txtUrl             = $window.FindName('TxtUrl')
$radStyleFill       = $window.FindName('RadStyleFill')
$radStyleFit        = $window.FindName('RadStyleFit')
$radStyleStretch    = $window.FindName('RadStyleStretch')
$radStyleTile       = $window.FindName('RadStyleTile')
$radStyleCenter     = $window.FindName('RadStyleCenter')
$radStyleSpan       = $window.FindName('RadStyleSpan')
$chkLockBackground  = $window.FindName('ChkLockBackground')

$cardLockScreen     = $window.FindName('CardLockScreen')
$chkSetLockScreen   = $window.FindName('ChkSetLockScreen')
$chkSameAsDesk      = $window.FindName('ChkSameAsDesk')
$pnlLockCustom      = $window.FindName('PnlLockCustom')
$radLockLocal       = $window.FindName('RadLockLocal')
$radLockUrl         = $window.FindName('RadLockUrl')
$pnlLockLocal       = $window.FindName('PnlLockLocal')
$pnlLockUrl         = $window.FindName('PnlLockUrl')
$txtLockLocalFile   = $window.FindName('TxtLockLocalFile')
$btnBrowseLockLocal = $window.FindName('BtnBrowseLockLocal')
$txtLockUrl         = $window.FindName('TxtLockUrl')
$chkLockLockScreen  = $window.FindName('ChkLockLockScreen')

$txtOutputDir       = $window.FindName('TxtOutputDir')
$btnBrowseOutput    = $window.FindName('BtnBrowseOutput')

$cardPreview        = $window.FindName('CardPreview')
$txtPreviewBadge    = $window.FindName('TxtPreviewBadge')
$radPrevDesktop     = $window.FindName('RadPrevDesktop')
$radPrevLock        = $window.FindName('RadPrevLock')
$imgPreview         = $window.FindName('ImgPreview')
$rectPreviewTile    = $window.FindName('RectPreviewTile')
$pnlPreviewEmpty    = $window.FindName('PnlPreviewEmpty')
$pnlOverlayDesktop  = $window.FindName('PnlOverlayDesktop')
$pnlOverlayLock     = $window.FindName('PnlOverlayLock')
$txtLockClock       = $window.FindName('TxtLockClock')
$txtLockDate        = $window.FindName('TxtLockDate')
$txtBuildLog        = $window.FindName('TxtBuildLog')

$txtFooterStatus    = $window.FindName('TxtFooterStatus')
$btnOpenOutput      = $window.FindName('BtnOpenOutput')
$btnGenerate        = $window.FindName('BtnGenerate')

$Script:LogTextBox = $txtBuildLog
$Script:FooterStatusText = $txtFooterStatus

# Default Output Directory
$defaultOutput = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
    Join-Path $PSScriptRoot 'Output'
} else {
    Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) 'Downloads\DesktopBackground_Output'
}
$txtOutputDir.Text = $defaultOutput

# Theme State + Painted Win11 Chrome Hook
$Script:IsDarkTheme = Get-SystemIsDarkTheme
$Script:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true

$window.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $Script:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $Script:IsSystemThemeAuto) {
        $window.Dispatcher.Invoke([Action]{
            $Script:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $true
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)
$window.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
})

$btnThemeToggle.Add_Click({
    $Script:IsSystemThemeAuto = $false
    $Script:IsDarkTheme = -not $Script:IsDarkTheme
    Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $txtThemeLabel -IsSystemAuto $false
})

# ---------------------------------------------------------------------------
# 6. Preview & Control State Synchronization
# ---------------------------------------------------------------------------
function Get-SelectedWallpaperStyle {
    [CmdletBinding()]
    param()
    if ($radStyleFit.IsChecked)     { return 'Fit' }
    if ($radStyleStretch.IsChecked) { return 'Stretch' }
    if ($radStyleTile.IsChecked)    { return 'Tile' }
    if ($radStyleCenter.IsChecked)  { return 'Center' }
    if ($radStyleSpan.IsChecked)    { return 'Span' }
    return 'Fill'
}

function New-FrozenBitmap {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ImagePath)
    $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
    $ms = New-Object System.IO.MemoryStream(, $bytes)
    try {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.StreamSource = $ms
        $bmp.EndInit()
        if ($bmp.CanFreeze) { $bmp.Freeze() }
        return $bmp
    } finally {
        $ms.Dispose()
    }
}

function Update-ScreenPreview {
    [CmdletBinding()]
    param()

    $isLockView = ($radPrevLock.IsChecked -eq $true)
    $selectedStyle = Get-SelectedWallpaperStyle
    $viewLabel = if ($isLockView) { 'Lock Screen' } else { '16:9' }
    $txtPreviewBadge.Text = "$selectedStyle · $viewLabel"

    $pnlOverlayDesktop.Visibility = [System.Windows.Visibility]::Collapsed
    $pnlOverlayLock.Visibility    = [System.Windows.Visibility]::Collapsed
    if ($isLockView) {
        $now = Get-Date
        $txtLockClock.Text = $now.ToString('HH:mm')
        $txtLockDate.Text  = $now.ToString('dddd, d MMMM')
    }

    $candidatePath = ''
    if ($isLockView) {
        if ($chkSameAsDesk.IsChecked -and $chkSetDesktop.IsChecked) {
            if ($radLocal.IsChecked) { $candidatePath = $txtLocalFile.Text.Trim() }
        } elseif ($chkSetLockScreen.IsChecked -and $radLockLocal.IsChecked) {
            $candidatePath = $txtLockLocalFile.Text.Trim()
        }
    } else {
        if ($chkSetDesktop.IsChecked -and $radLocal.IsChecked) {
            $candidatePath = $txtLocalFile.Text.Trim()
        }
    }

    $imgPreview.Source = $null
    $rectPreviewTile.Fill = $null
    $rectPreviewTile.Visibility = [System.Windows.Visibility]::Collapsed

    if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
        try {
            $bitmap = New-FrozenBitmap -ImagePath $candidatePath
            $pnlPreviewEmpty.Visibility = [System.Windows.Visibility]::Collapsed
            $pnlOverlayDesktop.Visibility = if ($isLockView) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
            $pnlOverlayLock.Visibility    = if ($isLockView) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

            if ($selectedStyle -eq 'Tile' -and -not $isLockView) {
                $tileBrush = New-Object System.Windows.Media.ImageBrush($bitmap)
                $tileBrush.TileMode = [System.Windows.Media.TileMode]::Tile
                $tileBrush.Viewport = New-Object System.Windows.Rect(0, 0, 0.45, 0.45)
                $tileBrush.ViewportUnits = [System.Windows.Media.BrushMappingMode]::RelativeToBoundingBox
                $tileBrush.Stretch = [System.Windows.Media.Stretch]::UniformToFill
                $rectPreviewTile.Fill = $tileBrush
                $rectPreviewTile.Visibility = [System.Windows.Visibility]::Visible
                $imgPreview.Visibility = [System.Windows.Visibility]::Collapsed
            } else {
                $imgPreview.Visibility = [System.Windows.Visibility]::Visible
                $imgPreview.Source = $bitmap
                $imgPreview.Margin = New-Object System.Windows.Thickness(0)
                if ($isLockView) {
                    $imgPreview.Stretch = [System.Windows.Media.Stretch]::UniformToFill
                } else {
                    switch ($selectedStyle) {
                        'Fit'     { $imgPreview.Stretch = [System.Windows.Media.Stretch]::Uniform }
                        'Stretch' { $imgPreview.Stretch = [System.Windows.Media.Stretch]::Fill }
                        'Center'  {
                            $imgPreview.Stretch = [System.Windows.Media.Stretch]::Uniform
                            $imgPreview.Margin = New-Object System.Windows.Thickness(48, 28, 48, 28)
                        }
                        default   { $imgPreview.Stretch = [System.Windows.Media.Stretch]::UniformToFill }
                    }
                }
            }
            return
        } catch {}
    }

    $pnlPreviewEmpty.Visibility = [System.Windows.Visibility]::Visible
}

function Update-ControlState {
    [CmdletBinding()]
    param()

    $deskEnabled = ($chkSetDesktop.IsChecked -eq $true)
    $deskIsLocal = ($radLocal.IsChecked -eq $true)

    $pnlDeskBody.IsEnabled = $deskEnabled
    $pnlDeskBody.Opacity   = if ($deskEnabled) { 1.0 } else { 0.45 }
    $radLocal.IsEnabled    = $deskEnabled
    $radUrl.IsEnabled      = $deskEnabled

    $pnlDeskLocal.Visibility = if ($deskIsLocal) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $pnlDeskUrl.Visibility   = if ($deskIsLocal) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }

    $lockEnabled = ($chkSetLockScreen.IsChecked -eq $true)
    if ($lockEnabled -and -not $deskEnabled -and $chkSameAsDesk.IsChecked) {
        $chkSameAsDesk.IsChecked = $false
    }
    $chkSameAsDesk.IsEnabled     = $lockEnabled -and $deskEnabled
    $chkLockLockScreen.IsEnabled = $lockEnabled

    $lockCustomVisible = $lockEnabled -and (-not $chkSameAsDesk.IsChecked)
    $pnlLockCustom.Visibility = if ($lockCustomVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $lockIsLocal = ($radLockLocal.IsChecked -eq $true)
    $pnlLockLocal.Visibility = if ($lockIsLocal) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $pnlLockUrl.Visibility   = if ($lockIsLocal) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }

    Update-ScreenPreview
}

# Wire Input State & Preview Events
$txtVersion.Add_TextChanged({
    $v = $txtVersion.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($v)) { $v = '2.0' }
    if ($v -notmatch '^[vV]') { $v = "v$v" }
    $txtVersionPill.Text = $v
})

foreach ($ctl in @($chkSetDesktop, $radLocal, $radUrl, $chkSetLockScreen, $chkSameAsDesk, $radLockLocal, $radLockUrl)) {
    $ctl.Add_Checked({ Update-ControlState })
    $ctl.Add_Unchecked({ Update-ControlState })
}
foreach ($styleRadio in @($radStyleFill, $radStyleFit, $radStyleStretch, $radStyleTile, $radStyleCenter, $radStyleSpan, $radPrevDesktop, $radPrevLock)) {
    $styleRadio.Add_Checked({ Update-ScreenPreview })
}
$txtLocalFile.Add_TextChanged({ Update-ScreenPreview })
$txtLockLocalFile.Add_TextChanged({ Update-ScreenPreview })

function Invoke-ImageDragDrop {
    param(
        [Parameter(Mandatory = $true)]$DropEvent,
        [ValidateSet('Auto', 'Desktop', 'Lock')][string]$Target = 'Auto'
    )
    if (-not $DropEvent.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
    $files = $DropEvent.Data.GetData([System.Windows.DataFormats]::FileDrop)
    if (-not $files -or $files.Count -eq 0) { return }
    $file = $files[0]
    $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
    if ($ext -notin @('.png', '.jpg', '.jpeg', '.bmp')) { return }

    $assignLock = ($Target -eq 'Lock') -or ($Target -eq 'Auto' -and $radPrevLock.IsChecked -and -not $chkSameAsDesk.IsChecked)
    if ($assignLock) {
        $chkSetLockScreen.IsChecked = $true
        $chkSameAsDesk.IsChecked = $false
        $radLockLocal.IsChecked = $true
        $txtLockLocalFile.Text = $file
        Write-BuildLog -Message "Loaded Lock Screen image via drag-and-drop: $file" -Level INFO
    } else {
        $chkSetDesktop.IsChecked = $true
        $radLocal.IsChecked = $true
        $txtLocalFile.Text = $file
        Write-BuildLog -Message "Loaded Desktop image via drag-and-drop: $file" -Level INFO
    }
    Update-ControlState
}

$cardDesktop.Add_Drop({ param($s, $e) Invoke-ImageDragDrop -DropEvent $e -Target 'Desktop' })
$cardLockScreen.Add_Drop({ param($s, $e) Invoke-ImageDragDrop -DropEvent $e -Target 'Lock' })
$cardPreview.Add_Drop({ param($s, $e) Invoke-ImageDragDrop -DropEvent $e -Target 'Auto' })

$btnBrowseLocal.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*'
    if ($dlg.ShowDialog($window) -eq $true) {
        $txtLocalFile.Text = $dlg.FileName
        Write-BuildLog -Message "Selected Desktop image: $($dlg.FileName)" -Level INFO
        Update-ScreenPreview
    }
})

$btnBrowseLockLocal.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*'
    if ($dlg.ShowDialog($window) -eq $true) {
        $txtLockLocalFile.Text = $dlg.FileName
        Write-BuildLog -Message "Selected Lock Screen image: $($dlg.FileName)" -Level INFO
        Update-ScreenPreview
    }
})

$btnBrowseOutput.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.SelectedPath = $txtOutputDir.Text
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $folderBrowser.SelectedPath
        Write-BuildLog -Message "Output directory changed to $($folderBrowser.SelectedPath)" -Level INFO
    }
})

$btnOpenOutput.Add_Click({
    $dir = $txtOutputDir.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        if (-not (Test-Path -Path $dir)) {
            try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch {}
        }
        if (Test-Path -Path $dir) {
            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$dir`""
        }
    }
})

# Initial Log Entry & State Sync
Update-ControlState
$initThemeName = if ($Script:IsDarkTheme) { 'Dark (Obsidian Teal)' } else { 'Light (Royal Blue)' }
Write-BuildLog -Message "Ready. Windows theme detected: $initThemeName." -Level INFO
Write-BuildLog -Message "Output target: $defaultOutput" -Level INFO

# ---------------------------------------------------------------------------
# 7. Intune Packaging Engine (Preserved 100% Output & 5-Step Guide Parity)
# ---------------------------------------------------------------------------
$btnGenerate.Add_Click({
    $btnGenerate.IsEnabled = $false
    $txtFooterStatus.Text = 'Starting packaging process...'
    $txtFooterStatus.Foreground = $window.Resources['BrushTextMuted']
    Write-BuildLog -Message 'Starting Intune Desktop Background packaging...' -Level INFO
    Invoke-UiRefresh

    $stagingDir = $null
    try {
        # 1. Validation
        $setDesktop      = ($chkSetDesktop.IsChecked -eq $true)
        $setLockScreen   = ($chkSetLockScreen.IsChecked -eq $true)
        $deskLocalActive = ($radLocal.IsChecked -eq $true)
        $deskUrlActive   = ($radUrl.IsChecked -eq $true)

        $lockIsSame      = ($chkSameAsDesk.IsChecked -eq $true) -and $setDesktop
        $lockIsLocalDiff = $setLockScreen -and (-not $lockIsSame) -and ($radLockLocal.IsChecked -eq $true)
        $lockIsUrlDiff   = $setLockScreen -and (-not $lockIsSame) -and ($radLockUrl.IsChecked -eq $true)
        $lockIsDiff      = $lockIsLocalDiff -or $lockIsUrlDiff

        $rawVersion = $txtVersion.Text.Trim()
        $cleanVersion = ($rawVersion -replace '[^a-zA-Z0-9\.\-]', '')
        if ([string]::IsNullOrWhiteSpace($cleanVersion)) { $cleanVersion = '2.0' }
        $versionTag = if ($cleanVersion -match '^[vV]') { $cleanVersion } else { "v$cleanVersion" }

        $selectedStyle = Get-SelectedWallpaperStyle
        switch ($selectedStyle) {
            'Fill'    { $wallpaperStyleVal = '10'; $tileWallpaperVal = '0' }
            'Fit'     { $wallpaperStyleVal = '6';  $tileWallpaperVal = '0' }
            'Stretch' { $wallpaperStyleVal = '2';  $tileWallpaperVal = '0' }
            'Center'  { $wallpaperStyleVal = '0';  $tileWallpaperVal = '0' }
            'Tile'    { $wallpaperStyleVal = '0';  $tileWallpaperVal = '1' }
            'Span'    { $wallpaperStyleVal = '22'; $tileWallpaperVal = '0' }
            default   { $wallpaperStyleVal = '10'; $tileWallpaperVal = '0' }
        }

        if (-not $setDesktop -and -not $setLockScreen) {
            [System.Windows.MessageBox]::Show('You must enable at least one option: Desktop Background or Lock Screen.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }

        if ($setDesktop) {
            if ($deskLocalActive -and (-not (Test-Path -LiteralPath $txtLocalFile.Text.Trim() -PathType Leaf))) {
                [System.Windows.MessageBox]::Show('Please select a valid local image file for the Desktop Background.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            if ($deskUrlActive -and [string]::IsNullOrWhiteSpace($txtUrl.Text)) {
                [System.Windows.MessageBox]::Show('Please enter a valid Desktop Background image URL.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
        }

        if ($setLockScreen -and $lockIsDiff) {
            if ($lockIsUrlDiff -and [string]::IsNullOrWhiteSpace($txtLockUrl.Text)) {
                [System.Windows.MessageBox]::Show('Please enter a valid URL for the Lock Screen.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
            if ($lockIsLocalDiff -and (-not (Test-Path -LiteralPath $txtLockLocalFile.Text.Trim() -PathType Leaf))) {
                [System.Windows.MessageBox]::Show('Please select a valid local image file for the Lock Screen.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
        }

        $outDir = $txtOutputDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($outDir)) {
            [System.Windows.MessageBox]::Show('Please specify an Output Directory.', 'Validation Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }
        if (-not (Test-Path -Path $outDir)) {
            try {
                New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            } catch {
                [System.Windows.MessageBox]::Show("Could not create output directory: $($_.Exception.Message)", 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
        }

        $isLocal = $true
        if ($setDesktop -and $deskUrlActive) { $isLocal = $false }
        if ($setLockScreen -and $lockIsUrlDiff) { $isLocal = $false }

        # 2. Check/Download IntuneWinAppUtil.exe
        $packerExe = Get-IntuneWinAppUtilPath -StatusTextBlock $txtFooterStatus
        if ([string]::IsNullOrWhiteSpace($packerExe) -or (-not (Test-Path -Path $packerExe -PathType Leaf))) {
            $txtFooterStatus.Text = 'Build aborted: IntuneWinAppUtil.exe is required.'
            $txtFooterStatus.Foreground = $window.Resources['BrushErr']
            Write-BuildLog -Message 'Build aborted: IntuneWinAppUtil.exe was not provided.' -Level WARN
            return
        }

        # 3. Setup Staging Directory
        $txtFooterStatus.Text = 'Staging scripts and wallpaper assets...'
        Write-BuildLog -Message "Using IntuneWinAppUtil: $packerExe" -Level INFO
        $stagingBase = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
            $PSScriptRoot
        } else {
            $env:TEMP
        }
        $stagingDir = Join-Path $stagingBase "Staging_IntuneBackground_$([guid]::NewGuid().ToString().Substring(0,8))"
        if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

        # 4. Generate Install & Payload Scripts
        $installScriptPath = Join-Path $stagingDir 'Install.ps1'
        $installContent = ''

        if ($isLocal) {
            $fileName = ''
            if ($setDesktop) {
                $fileName = Split-Path $txtLocalFile.Text.Trim() -Leaf
                $destImg = Join-Path $stagingDir "desktop_$fileName"
                Copy-Item $txtLocalFile.Text.Trim() -Destination $destImg -Force
            }

            $lockFileName = ''
            if ($setLockScreen) {
                if ($lockIsSame) {
                    $lockFileName = "desktop_$fileName"
                } else {
                    $lockOriginalName = Split-Path $txtLockLocalFile.Text.Trim() -Leaf
                    $lockFileName = "lock_$lockOriginalName"
                    $destLockImg = Join-Path $stagingDir $lockFileName
                    Copy-Item $txtLockLocalFile.Text.Trim() -Destination $destLockImg -Force
                }
            }

            if ($setDesktop) {
                $applyScriptContent = @"
`$ErrorActionPreference = 'Continue'
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\ApplyBackground.log`" -Append
`$destDesktop = "`$env:ProgramData\CorporateBackground\desktop_$fileName"

# Set Wallpaper Position / Style in user registry ($selectedStyle)
`$desktopReg = `"HKCU:\Control Panel\Desktop`"
Set-ItemProperty -Path `$desktopReg -Name `"WallpaperStyle`" -Value "$wallpaperStyleVal" -Force
Set-ItemProperty -Path `$desktopReg -Name `"TileWallpaper`" -Value "$tileWallpaperVal" -Force

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@

[Wallpaper]::SystemParametersInfo(0x0014, 0, `$destDesktop, 0x01 -bOr 0x02)
Stop-Transcript
"@
                Set-Content -Path (Join-Path $stagingDir 'ApplyBackground.ps1') -Value $applyScriptContent -Force
            }

            $installContent = @"
`$ErrorActionPreference = 'Continue'
`$targetPath = "`$env:ProgramData\CorporateBackground"
if (!(Test-Path `$targetPath)) { New-Item -ItemType Directory -Path `$targetPath -Force }
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\Install.log`" -Force

# Write version marker for Intune detection and update triggers
Set-Content -Path "`$targetPath\version.txt" -Value "$cleanVersion" -Force
"@
            if ($setDesktop) {
                $installContent += @"

Copy-Item ".\desktop_$fileName" -Destination "`$targetPath\desktop_$fileName" -Force
Copy-Item ".\ApplyBackground.ps1" -Destination "`$targetPath\ApplyBackground.ps1" -Force
"@
            }
            if ($setLockScreen) {
                if ($lockIsDiff) {
                    $installContent += @"

Copy-Item ".\$lockFileName" -Destination "`$targetPath\$lockFileName" -Force
`$destLock = "`$targetPath\$lockFileName"
"@
                } else {
                    $installContent += @"

`$destLock = "`$targetPath\desktop_$fileName"
"@
                }
                $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"LockScreenImage`" -Value `$destLock -Force
"@
            }

            if ($setDesktop) {
                $installContent += @"

`$taskName = `"ApplyCorporateBackground`"
`$actionScriptPath = "`$targetPath\ApplyBackground.ps1"
`$action = New-ScheduledTaskAction -Execute `"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`" -Argument `"-WindowStyle Hidden -ExecutionPolicy Bypass -Command ``"Get-Content '`$actionScriptPath' -Raw | Invoke-Expression``""`
`$trigger = New-ScheduledTaskTrigger -AtLogon
`$principal = New-ScheduledTaskPrincipal -GroupId `"BUILTIN\Users`"
`$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName `$taskName -Action `$action -Trigger `$trigger -Principal `$principal -Settings `$settings -Force

Start-ScheduledTask -TaskName `$taskName
"@
            }

            if ($setDesktop) {
                if ($chkLockBackground.IsChecked) {
                    $installContent += @"

Start-Sleep -Seconds 5
`$ActiveDesktopPath = `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`"
if (!(Test-Path `$ActiveDesktopPath)) { New-Item -Path `$ActiveDesktopPath -Force | Out-Null }
Set-ItemProperty -Path `$ActiveDesktopPath -Name `"NoChangingWallPaper`" -Value 1 -Force
"@
                } else {
                    $installContent += @"

# Ensure wallpaper lock is removed if previously set
try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
try {
    Get-ChildItem Registry::HKEY_USERS -ErrorAction Stop | ForEach-Object {
        `$sid = `$_.PSChildName
        if (`$sid -notmatch '_Classes$' -and `$sid -match '^S-1-5-21-') {
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
        }
    }
} catch {}
"@
                }
            }

            if ($setLockScreen) {
                if ($chkLockLockScreen.IsChecked) {
                    $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"NoChangingLockScreen`" -Value 1 -Force
"@
                } else {
                    $installContent += @"

# Ensure lock screen lock is removed if previously set
try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
try {
    Get-ChildItem Registry::HKEY_USERS -ErrorAction Stop | ForEach-Object {
        `$sid = `$_.PSChildName
        if (`$sid -notmatch '_Classes$' -and `$sid -match '^S-1-5-21-') {
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
        }
    }
} catch {}
"@
                }
            }

            $installContent += @"

Stop-Transcript
"@
            Set-Content -Path $installScriptPath -Value $installContent -Force
        } else {
            $desktopUrl = if ($setDesktop -and $deskUrlActive) { $txtUrl.Text.Trim() } else { '' }
            $lockUrl = if ($setLockScreen) {
                if ($lockIsSame) { $desktopUrl } elseif ($lockIsUrlDiff) { $txtLockUrl.Text.Trim() } else { '' }
            } else { '' }

            $desktopExt = '.png'
            if ($setDesktop -and $deskUrlActive) {
                $desktopExt = [System.IO.Path]::GetExtension($desktopUrl)
                if ($desktopExt -notmatch '\.(jpg|jpeg|png|bmp)') { $desktopExt = '.png' }
            }

            $lockExt = '.png'
            if ($setLockScreen) {
                if ($lockIsSame) {
                    $lockExt = $desktopExt
                } elseif ($lockIsUrlDiff) {
                    $lockExt = [System.IO.Path]::GetExtension($lockUrl)
                    if ($lockExt -notmatch '\.(jpg|jpeg|png|bmp)') { $lockExt = '.png' }
                }
            }

            $fileName = ''
            if ($setDesktop -and $deskLocalActive) {
                $fileName = Split-Path $txtLocalFile.Text.Trim() -Leaf
                $destImg = Join-Path $stagingDir "desktop_$fileName"
                Copy-Item $txtLocalFile.Text.Trim() -Destination $destImg -Force
            }

            $lockFileName = ''
            $lockOriginalName = ''
            if ($setLockScreen -and $lockIsLocalDiff) {
                $lockOriginalName = Split-Path $txtLockLocalFile.Text.Trim() -Leaf
                $lockFileName = "lock_$lockOriginalName"
                $destLockImg = Join-Path $stagingDir $lockFileName
                Copy-Item $txtLockLocalFile.Text.Trim() -Destination $destLockImg -Force
            }

            $destDesktopPath = ''
            if ($setDesktop) {
                if ($deskLocalActive) {
                    $destDesktopPath = "`$targetPath\desktop_$fileName"
                } else {
                    $destDesktopPath = "`$targetPath\desktop_bg$desktopExt"
                }
            }

            $destLockPath = ''
            if ($setLockScreen) {
                if ($lockIsSame) {
                    if ($deskLocalActive) {
                        $destLockPath = "`$targetPath\desktop_$fileName"
                    } else {
                        $destLockPath = "`$targetPath\desktop_bg$desktopExt"
                    }
                } elseif ($lockIsLocalDiff) {
                    $destLockPath = "`$targetPath\lock_$lockOriginalName"
                } else {
                    $destLockPath = "`$targetPath\lock_bg$lockExt"
                }
            }

            $dlContent = @"
`$ErrorActionPreference = 'Continue'
`$targetPath = "`$env:ProgramData\CorporateBackground"
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\DownloadBackgrounds.log`" -Append

# Sleep for a random duration (0 to 5 minutes) to stagger requests and prevent corporate IP rate-limiting
`$randomSleep = Get-Random -Minimum 0 -Maximum 300
Write-Output "Staggering download request: Sleeping for `$randomSleep seconds..."
Start-Sleep -Seconds `$randomSleep

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
"@

            if ($setDesktop -and $deskUrlActive) {
                $dlContent += @"

`$localFile = "$destDesktopPath"
`$needsDownload = `$true

if (Test-Path `$localFile) {
    try {
        `$headers = Invoke-WebRequest -Uri '$desktopUrl' -Method Head -UseBasicParsing -ErrorAction Stop
        `$remoteLength = `$headers.Headers['Content-Length']
        `$localLength = (Get-Item `$localFile).Length

        if (`$remoteLength -eq `$localLength) {
            Write-Output "Remote file size (`$remoteLength) matches local file size. Skipping download."
            `$needsDownload = `$false
        }
    } catch {
        Write-Output "Could not verify remote file size: `$($_.Exception.Message). Proceeding with download."
    }
}

if (`$needsDownload) {
    try {
        Write-Output `"Attempting to download desktop background from: $desktopUrl`"
        Invoke-WebRequest -Uri '$desktopUrl' -OutFile "$destDesktopPath" -UseBasicParsing
        Write-Output `"Desktop background downloaded successfully to $destDesktopPath`"
    } catch {
        Write-Output `"ERROR: Failed to download desktop background.`"
        Write-Output `"Exception: `$(`$_.Exception.Message)`"
    }
}
"@
            }

            if ($setLockScreen) {
                if ($lockIsUrlDiff) {
                    $dlContent += @"

`$localLockFile = "$destLockPath"
`$needsLockDownload = `$true

if (Test-Path `$localLockFile) {
    try {
        `$headers = Invoke-WebRequest -Uri '$lockUrl' -Method Head -UseBasicParsing -ErrorAction Stop
        `$remoteLength = `$headers.Headers['Content-Length']
        `$localLength = (Get-Item `$localLockFile).Length

        if (`$remoteLength -eq `$localLength) {
            Write-Output "Remote lock screen size (`$remoteLength) matches local file size. Skipping download."
            `$needsLockDownload = `$false
        }
    } catch {
        Write-Output "Could not verify remote lock screen size: `$($_.Exception.Message). Proceeding with download."
    }
}

if (`$needsLockDownload) {
    try {
        Write-Output `"Attempting to download lock screen from: $lockUrl`"
        Invoke-WebRequest -Uri '$lockUrl' -OutFile "$destLockPath" -UseBasicParsing
        Write-Output `"Lock screen downloaded successfully to $destLockPath`"
    } catch {
        Write-Output `"ERROR: Failed to download lock screen.`"
        Write-Output `"Exception: `$(`$_.Exception.Message)`"
    }
}
"@
                } elseif ($lockIsSame -and $deskUrlActive) {
                    $dlContent += @"
if (`$needsDownload) {
    Copy-Item "$destDesktopPath" -Destination "$destLockPath" -Force
}
"@
                }
            }

            if ($setLockScreen) {
                $dlContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"LockScreenImage`" -Value "$destLockPath" -Force
"@
            }

            if ($setDesktop -and $chkLockBackground.IsChecked) {
                $dlContent += @"

try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop } catch {}
"@
            }

            if ($setDesktop) {
                $dlContent += @"

try { Start-ScheduledTask -TaskName 'ApplyCorporateBackground' -ErrorAction Stop } catch {}
"@
            }

            if ($setDesktop -and $chkLockBackground.IsChecked) {
                $dlContent += @"

Start-Sleep -Seconds 5
`$ActiveDesktopPath = `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`"
if (!(Test-Path `$ActiveDesktopPath)) { New-Item -Path `$ActiveDesktopPath -Force | Out-Null }
Set-ItemProperty -Path `$ActiveDesktopPath -Name `"NoChangingWallPaper`" -Value 1 -Force
"@
            }

            $dlContent += @"

Stop-Transcript
"@
            Set-Content -Path (Join-Path $stagingDir 'DownloadBackgrounds.ps1') -Value $dlContent -Force

            if ($setDesktop) {
                $applyScriptContent = @"
`$ErrorActionPreference = 'Continue'
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\ApplyBackground.log`" -Append
`$destDesktop = "$destDesktopPath"

# Set Wallpaper Position / Style in user registry ($selectedStyle)
`$desktopReg = `"HKCU:\Control Panel\Desktop`"
Set-ItemProperty -Path `$desktopReg -Name `"WallpaperStyle`" -Value "$wallpaperStyleVal" -Force
Set-ItemProperty -Path `$desktopReg -Name `"TileWallpaper`" -Value "$tileWallpaperVal" -Force

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
'@

[Wallpaper]::SystemParametersInfo(0x0014, 0, `$destDesktop, 0x01 -bOr 0x02)
Stop-Transcript
"@
                Set-Content -Path (Join-Path $stagingDir 'ApplyBackground.ps1') -Value $applyScriptContent -Force
            }

            $installContent = @"
`$ErrorActionPreference = 'Continue'
`$targetPath = "`$env:ProgramData\CorporateBackground"
if (!(Test-Path `$targetPath)) { New-Item -ItemType Directory -Path `$targetPath -Force }
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\Install.log`" -Force

# Write version marker for Intune detection and update triggers
Set-Content -Path "`$targetPath\version.txt" -Value "$cleanVersion" -Force

# Performing initial downloads...
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
"@
            if ($setDesktop) {
                if ($deskLocalActive) {
                    $installContent += @"

Copy-Item ".\desktop_$fileName" -Destination "$destDesktopPath" -Force
"@
                } else {
                    $installContent += @"

try {
    Write-Output "Downloading desktop background from: $desktopUrl"
    Invoke-WebRequest -Uri '$desktopUrl' -OutFile "$destDesktopPath" -UseBasicParsing
    Write-Output "Desktop background downloaded successfully."
} catch {
    Write-Error "ERROR: Failed to download initial desktop background: `$($_.Exception.Message)"
    Stop-Transcript
    exit 1
}
"@
                }
            }

            if ($setLockScreen) {
                if ($lockIsLocalDiff) {
                    $installContent += @"

Copy-Item ".\$lockFileName" -Destination "$destLockPath" -Force
"@
                } elseif ($lockIsUrlDiff) {
                    $installContent += @"

try {
    Write-Output "Downloading lock screen from: $lockUrl"
    Invoke-WebRequest -Uri '$lockUrl' -OutFile "$destLockPath" -UseBasicParsing
    Write-Output "Lock screen downloaded successfully."
} catch {
    Write-Error "ERROR: Failed to download initial lock screen: `$($_.Exception.Message)"
    Stop-Transcript
    exit 1
}
"@
                } elseif ($lockIsSame) {
                    $installContent += @"

Copy-Item "$destDesktopPath" -Destination "$destLockPath" -Force
"@
                }

                $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"LockScreenImage`" -Value "$destLockPath" -Force
"@
            }

            if ($setDesktop) {
                $installContent += @"

Copy-Item ".\ApplyBackground.ps1" -Destination "`$targetPath\ApplyBackground.ps1" -Force
"@
            }

            $installContent += @"

Copy-Item ".\DownloadBackgrounds.ps1" -Destination "`$targetPath\DownloadBackgrounds.ps1" -Force

`$taskNameDL = `"CorporateBackgroundDownloader`"
`$dlScriptPath = "`$targetPath\DownloadBackgrounds.ps1"
`$dlAction = New-ScheduledTaskAction -Execute `"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`" -Argument `"-WindowStyle Hidden -ExecutionPolicy Bypass -Command ``"Get-Content '`$dlScriptPath' -Raw | Invoke-Expression``""`
`$dlTrigger1 = New-ScheduledTaskTrigger -Daily -At 10:00AM
`$dlTrigger2 = New-ScheduledTaskTrigger -AtLogon
`$dlPrincipal = New-ScheduledTaskPrincipal -UserId `"SYSTEM`" -RunLevel Highest
`$dlSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName `$taskNameDL -Action `$dlAction -Trigger `$dlTrigger1, `$dlTrigger2 -Principal `$dlPrincipal -Settings `$dlSettings -Force
"@

            if ($setDesktop) {
                $installContent += @"

`$taskNameApply = `"ApplyCorporateBackground`"
`$applyScriptPath = "`$targetPath\ApplyBackground.ps1"
`$applyAction = New-ScheduledTaskAction -Execute `"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`" -Argument `"-WindowStyle Hidden -ExecutionPolicy Bypass -Command ``"Get-Content '`$applyScriptPath' -Raw | Invoke-Expression``""`
`$applyTrigger = New-ScheduledTaskTrigger -AtLogon
`$applyPrincipal = New-ScheduledTaskPrincipal -GroupId `"BUILTIN\Users`"
`$applySettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName `$taskNameApply -Action `$applyAction -Trigger `$applyTrigger -Principal `$applyPrincipal -Settings `$applySettings -Force

Start-Sleep -Seconds 2
Write-Output `"Applying background to current user...`"
Start-ScheduledTask -TaskName `$taskNameApply
"@
            }

            if ($setDesktop) {
                if ($chkLockBackground.IsChecked) {
                    $installContent += @"

Start-Sleep -Seconds 5
`$ActiveDesktopPath = `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`"
if (!(Test-Path `$ActiveDesktopPath)) { New-Item -Path `$ActiveDesktopPath -Force | Out-Null }
Set-ItemProperty -Path `$ActiveDesktopPath -Name `"NoChangingWallPaper`" -Value 1 -Force
"@
                } else {
                    $installContent += @"

# Ensure wallpaper lock is removed if previously set
try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
try {
    Get-ChildItem Registry::HKEY_USERS -ErrorAction Stop | ForEach-Object {
        `$sid = `$_.PSChildName
        if (`$sid -notmatch '_Classes$' -and `$sid -match '^S-1-5-21-') {
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
        }
    }
} catch {}
"@
                }
            }

            if ($setLockScreen) {
                if ($chkLockLockScreen.IsChecked) {
                    $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"NoChangingLockScreen`" -Value 1 -Force
"@
                } else {
                    $installContent += @"

# Ensure lock screen lock is removed if previously set
try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
try {
    Get-ChildItem Registry::HKEY_USERS -ErrorAction Stop | ForEach-Object {
        `$sid = `$_.PSChildName
        if (`$sid -notmatch '_Classes$' -and `$sid -match '^S-1-5-21-') {
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
        }
    }
} catch {}
"@
                }
            }

            $installContent += @"

Stop-Transcript
"@
            Set-Content -Path $installScriptPath -Value $installContent -Force
        }

        # Create Uninstall.ps1
        $uninstallContent = @"
`$ErrorActionPreference = 'Continue'

try {
    `$logDir = `"C:\Logs\DesktopBackground`"
    if (!(Test-Path `$logDir)) { New-Item -ItemType Directory -Path `$logDir -Force -ErrorAction Stop | Out-Null }
    Start-Transcript -Path "`$logDir\Uninstall.log`" -Append -Force -ErrorAction Stop
} catch {}

Write-Output "Starting Intune Desktop Background uninstallation..."

# 1. Stop and Unregister Scheduled Tasks
`$tasks = @('ApplyCorporateBackground', 'CorporateBackgroundDownloader')
foreach (`$t in `$tasks) {
    try { Stop-ScheduledTask -TaskName `$t -ErrorAction Stop | Out-Null } catch {}
    try {
        Unregister-ScheduledTask -TaskName `$t -Confirm:`$false -ErrorAction Stop | Out-Null
        Write-Output "Stopped and unregistered task: `$t"
    } catch {}
}

# 2. Remove Lock Policies from HKLM
`$hklmPaths = @(
    `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`",
    `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
)
foreach (`$p in `$hklmPaths) {
    if (Test-Path `$p) {
        try { Remove-ItemProperty -Path `$p -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
        try { Remove-ItemProperty -Path `$p -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
        try { Remove-ItemProperty -Path `$p -Name `"LockScreenImage`" -ErrorAction Stop | Out-Null } catch {}
    }
}

# 3. Remove Lock Policies from loaded User Hives (HKEY_USERS) and HKCU
try {
    Get-ChildItem Registry::HKEY_USERS -ErrorAction Stop | ForEach-Object {
        `$sid = `$_.PSChildName
        if (`$sid -notmatch '_Classes$' -and `$sid -match '^S-1-5-21-') {
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
        }
    }
} catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop | Out-Null } catch {}
try { Remove-ItemProperty -Path `"HKCU:\Software\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}

# 4. Clean up staged files (remove scripts first so detection rule immediately fails)
`$targetPath = "`$env:ProgramData\CorporateBackground"
if (Test-Path `$targetPath) {
    try { Get-ChildItem -Path `$targetPath -Filter "*.ps1" -ErrorAction Stop | ForEach-Object { try { Remove-Item -LiteralPath `$_.FullName -Force -ErrorAction Stop } catch {} } } catch {}
    try { Get-ChildItem -Path `$targetPath -Filter "*.txt" -ErrorAction Stop | ForEach-Object { try { Remove-Item -LiteralPath `$_.FullName -Force -ErrorAction Stop } catch {} } } catch {}
    try { Remove-Item -Path `$targetPath -Recurse -Force -ErrorAction Stop } catch {}
}

Write-Output "Intune Desktop Background uninstallation completed successfully."
try { Stop-Transcript -ErrorAction Stop } catch {}

exit 0
"@
        Set-Content -Path (Join-Path $stagingDir 'Uninstall.ps1') -Value $uninstallContent -Force

        # 5. Run IntuneWinAppUtil.exe
        $txtFooterStatus.Text = 'Packaging with IntuneWinAppUtil.exe...'
        Write-BuildLog -Message 'Running IntuneWinAppUtil.exe to compile Install.intunewin...' -Level INFO
        $process = Start-Process -FilePath $packerExe -ArgumentList "-c `"$stagingDir`"", '-s "Install.ps1"', "-o `"$outDir`"", '-q' -NoNewWindow -PassThru -Wait
        if ($process.ExitCode -ne 0) {
            throw "IntuneWinAppUtil failed to package the application. Exit code: $($process.ExitCode)"
        }

        $dateStr = Get-Date -Format 'ddMMyy'
        $newWimName = "DesktopBackground_${versionTag}_$dateStr.intunewin"
        $wimPath = Join-Path $outDir 'Install.intunewin'
        $newWimPath = Join-Path $outDir $newWimName
        Start-Sleep -Milliseconds 500

        if (Test-Path -LiteralPath $wimPath) {
            if (Test-Path -LiteralPath $newWimPath) { try { Remove-Item -LiteralPath $newWimPath -Force -ErrorAction Stop } catch {} }
            $renamed = $false
            for ($i = 1; $i -le 5; $i++) {
                try {
                    Rename-Item -LiteralPath $wimPath -NewName $newWimName -Force -ErrorAction Stop
                    $renamed = $true
                    break
                } catch {
                    Start-Sleep -Seconds 1
                }
            }
            if (-not $renamed) {
                throw "Failed to rename output package (file is locked by another process like OneDrive). You can manually rename 'Install.intunewin' to '$newWimName'."
            }
        } else {
            throw "Packaging failed: Could not find 'Install.intunewin' at expected path: '$wimPath'."
        }

        # 6. Generate DetectionScript.ps1 & 5-Step Intune InstallCommands.txt Guide
        if ($isLocal) {
            if ($setDesktop) {
                $detectionContent = @"
`$targetPath = "`$env:ProgramData\CorporateBackground"
`$logDir = `"C:\Logs\DesktopBackground`"
if (!(Test-Path `$logDir)) { try { New-Item -ItemType Directory -Path `$logDir -Force -ErrorAction Stop | Out-Null } catch {} }
`$timeStamp = Get-Date -Format `"dd/MM/yy HH:mm:ss`"
`$versionFile = "`$targetPath\version.txt"
`$rawVer = if (Test-Path `$versionFile) { try { Get-Content `$versionFile -Raw -ErrorAction Stop } catch { `$null } } else { `$null }
`$versionMatch = (`$null -ne `$rawVer) -and (`$rawVer.Trim() -eq "$cleanVersion")
`$schedTask = try { Get-ScheduledTask -TaskName 'ApplyCorporateBackground' -ErrorAction Stop } catch { `$null }
if ((Test-Path "`$targetPath\ApplyBackground.ps1") -and `$schedTask -and `$versionMatch) {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: INSTALLED (v$cleanVersion)" -ErrorAction Stop } catch {}
    Write-Output "Detected"
    exit 0
} else {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: NOT INSTALLED" -ErrorAction Stop } catch {}
    exit 1
}
"@
            } else {
                $detectionContent = @"
`$targetPath = "`$env:ProgramData\CorporateBackground"
`$logDir = `"C:\Logs\DesktopBackground`"
if (!(Test-Path `$logDir)) { try { New-Item -ItemType Directory -Path `$logDir -Force -ErrorAction Stop | Out-Null } catch {} }
`$timeStamp = Get-Date -Format `"dd/MM/yy HH:mm:ss`"
`$versionFile = "`$targetPath\version.txt"
`$rawVer = if (Test-Path `$versionFile) { try { Get-Content `$versionFile -Raw -ErrorAction Stop } catch { `$null } } else { `$null }
`$versionMatch = (`$null -ne `$rawVer) -and (`$rawVer.Trim() -eq "$cleanVersion")
`$regVal = try { (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name "LockScreenImage" -ErrorAction Stop).LockScreenImage } catch { `$null }
if (`$regVal -and (Test-Path `$regVal) -and `$versionMatch) {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: INSTALLED (v$cleanVersion)" -ErrorAction Stop } catch {}
    Write-Output "Detected"
    exit 0
} else {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: NOT INSTALLED" -ErrorAction Stop } catch {}
    exit 1
}
"@
            }
        } else {
            $detectionContent = @"
`$targetPath = "`$env:ProgramData\CorporateBackground"
`$logDir = `"C:\Logs\DesktopBackground`"
if (!(Test-Path `$logDir)) { try { New-Item -ItemType Directory -Path `$logDir -Force -ErrorAction Stop | Out-Null } catch {} }
`$timeStamp = Get-Date -Format `"dd/MM/yy HH:mm:ss`"
`$versionFile = "`$targetPath\version.txt"
`$rawVer = if (Test-Path `$versionFile) { try { Get-Content `$versionFile -Raw -ErrorAction Stop } catch { `$null } } else { `$null }
`$versionMatch = (`$null -ne `$rawVer) -and (`$rawVer.Trim() -eq "$cleanVersion")
`$schedTask = try { Get-ScheduledTask -TaskName 'CorporateBackgroundDownloader' -ErrorAction Stop } catch { `$null }
if ((Test-Path "`$targetPath\DownloadBackgrounds.ps1") -and `$schedTask -and `$versionMatch) {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: INSTALLED (v$cleanVersion)" -ErrorAction Stop } catch {}
    Write-Output "Detected"
    exit 0
} else {
    try { Add-Content -Path "`$logDir\Detection.log" -Value "[`$timeStamp] [DETECTION] STATUS: NOT INSTALLED" -ErrorAction Stop } catch {}
    exit 1
}
"@
        }

        $detectionFile = Join-Path $outDir 'DetectionScript.ps1'
        Set-Content -Path $detectionFile -Value $detectionContent -Force

        $generatedTimestamp = Get-Date -Format 'dd/MM/yy HH:mm:ss'
        $createdDateStr = Get-Date -Format 'dd/MM/yy'
        $lockBgChecked = ($chkLockBackground.IsChecked -eq $true)
        $lockScrChecked = ($chkLockLockScreen.IsChecked -eq $true)
        $desktopSummary = if ($setDesktop) { "Enabled (Style: $selectedStyle | Lock: $lockBgChecked)" } else { 'Disabled' }
        $lockSummary    = if ($setLockScreen) { "Enabled (Lock: $lockScrChecked)" } else { 'Disabled' }

        $commandsContent = @"
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : $generatedTimestamp
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
$newWimName


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
Corporate Desktop and Lock Screen Background ($versionTag)

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
Configures corporate Desktop Background (Wallpaper Style: $selectedStyle)
and Lock Screen image automatically for all users on the device.
All runtime logs are saved to C:\Logs\DesktopBackground\

PACKAGE DETAILS:
- Desktop Background : $desktopSummary
- Lock Screen        : $lockSummary
- Author             : JP
- App Version        : $cleanVersion
- Created            : $createdDateStr

WORKSTATION LOG LOCATION:
C:\Logs\DesktopBackground\

LOG FILES IN THIS FOLDER:
- Install Log   : C:\Logs\DesktopBackground\Install.log
- Apply Log     : C:\Logs\DesktopBackground\ApplyBackground.log
- Detection Log : C:\Logs\DesktopBackground\Detection.log
- Uninstall Log : C:\Logs\DesktopBackground\Uninstall.log
======================================================================

Publisher:
Pheneox

App Version:
$cleanVersion


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall.ps1

Install behavior:
System

Device restart behavior:
No specific action


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
DetectionScript.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No


========================================================================
ADDITIONAL NOTES
========================================================================
* Note for updating existing Intune apps: When replacing your .intunewin
  package to push a new image or changes to existing devices, always upload
  the new DetectionScript.ps1 as well. Intune will compare the version tag
  ($cleanVersion), detect that existing devices need an update, and deploy
  the new package seamlessly.
========================================================================
"@
        $commandsFile = Join-Path $outDir 'InstallCommands.txt'
        Set-Content -Path $commandsFile -Value $commandsContent -Force

        # 7. Cleanup
        if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }

        $txtFooterStatus.Text = "Package ready: $newWimName"
        $txtFooterStatus.Foreground = $window.Resources['BrushOk']
        Write-BuildLog -Message "Generated $newWimName, DetectionScript.ps1, and InstallCommands.txt in $outDir" -Level SUCCESS
        [System.Windows.MessageBox]::Show("Packaging Complete! Check the output directory for your $newWimName, DetectionScript.ps1, and InstallCommands.txt.", 'Success', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    } catch {
        if ($stagingDir -and (Test-Path $stagingDir)) { try { Remove-Item $stagingDir -Recurse -Force -ErrorAction Stop } catch {} }
        $txtFooterStatus.Text = "Error: $($_.Exception.Message)"
        $txtFooterStatus.Foreground = $window.Resources['BrushErr']
        Write-BuildLog -Message "Packaging failed: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("An error occurred: $($_.Exception.Message)", 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    } finally {
        $btnGenerate.IsEnabled = $true
    }
})

[void]$window.ShowDialog()
