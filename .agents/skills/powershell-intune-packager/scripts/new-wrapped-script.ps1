<#
.SYNOPSIS
    Scaffolds a new wrapped-folder PowerShell tool matching repository standards.
.DESCRIPTION
    Creates Scripts/<Category>/<ToolName>/ with <ToolName>.ps1, README.md, .gitignore,
    and an optional Start-Gui.bat launcher.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Azure', 'Clean Up Tools', 'Configuration', 'Datto', 'Intune', 'Kiosk', 'Microsoft 365', 'Networking', 'Sample', 'Software')]
    [string]$Category,

    [Parameter(Mandatory = $true)]
    [string]$ToolName,

    [Parameter(Mandatory = $false)]
    [string]$Synopsis = "Automated PowerShell utility for $ToolName.",

    [Parameter(Mandatory = $false)]
    [switch]$Gui,

    [Parameter(Mandatory = $false)]
    [switch]$Intune
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..\..\..")
$targetDir = Join-Path $repoRoot "Scripts\$Category\$ToolName"

if (-not (Test-Path -Path $targetDir)) {
    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
}

$scriptPath = Join-Path $targetDir "$ToolName.ps1"
$readmePath = Join-Path $targetDir "README.md"
$gitignorePath = Join-Path $targetDir ".gitignore"
$dateStr = Get-Date -Format "dd/MM/yyyy"
$shortDateStr = Get-Date -Format "dd/MM/yy"

$staBlock = if ($Gui) {
@"

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    `$relaunchPath = if (`$PSScriptRoot) { Join-Path `$PSScriptRoot `$MyInvocation.MyCommand.Name } else { `$PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"`$relaunchPath`""
    exit
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
"@
} else { "" }

$guiHelperBlock = if ($Gui) {
@"

if (-not ('Win11ChromeHelper' -as [type])) {
    `$chromeCSharp = @'
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
    try { Add-Type -TypeDefinition `$chromeCSharp -Language CSharp -ErrorAction Stop } catch {}
}

function Get-SystemIsDarkTheme {
    [CmdletBinding()]
    param()
    try {
        `$regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        if (Test-Path -Path `$regPath) {
            `$val = Get-ItemPropertyValue -Path `$regPath -Name 'AppsUseLightTheme' -ErrorAction Stop
            return ([int]`$val -eq 0)
        }
    } catch {}
    return `$true
}

function Set-StandardGuiTheme {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = `$true)]
        [System.Windows.Window]`$Window,

        [Parameter(Mandatory = `$true)]
        [bool]`$IsDark,

        [System.Windows.Controls.TextBlock]`$ThemeLabel = `$null,
        [bool]`$IsSystemAuto = `$true
    )

    `$conv = New-Object System.Windows.Media.BrushConverter
    if (`$IsDark) {
        `$canvasBrush = New-Object System.Windows.Media.RadialGradientBrush
        `$canvasBrush.Center = New-Object System.Windows.Point(0.38, 0.0)
        `$canvasBrush.GradientOrigin = New-Object System.Windows.Point(0.38, 0.0)
        `$canvasBrush.RadiusX = 0.95
        `$canvasBrush.RadiusY = 0.9
        `$canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#0F1F3D')), 0.0))
        `$canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#091122')), 0.52))
        `$canvasBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop ([System.Windows.Media.ColorConverter]::ConvertFromString('#05070C')), 1.0))

        `$map = @{
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
        `$canvasBrush = `$conv.ConvertFromString('#F8F9FF')
        `$map = @{
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

    if (`$canvasBrush.CanFreeze) { `$canvasBrush.Freeze() }
    `$Window.Background = `$canvasBrush
    `$Window.Resources['BrushCanvas'] = `$canvasBrush
    foreach (`$key in `$map.Keys) {
        `$b = `$conv.ConvertFromString(`$map[`$key])
        if (`$b.CanFreeze) { `$b.Freeze() }
        `$Window.Resources[`$key] = `$b
    }

    if (`$null -ne `$ThemeLabel) {
        `$modeTag = if (`$IsSystemAuto) { 'System' } else { 'Manual' }
        `$toneTag = if (`$IsDark) { 'Dark' } else { 'Light' }
        `$ThemeLabel.Text = "Theme: `$modeTag (`$toneTag)"
    }

    try {
        `$helper = New-Object System.Windows.Interop.WindowInteropHelper(`$Window)
        if (`$helper.Handle -ne [IntPtr]::Zero -and ('Win11ChromeHelper' -as [type])) {
            [Win11ChromeHelper]::ApplyWindowChrome(`$helper.Handle, `$IsDark)
        }
    } catch {}
}

function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = `$true)]
        [System.Windows.Window]`$Window,

        [switch]`$PreserveSize
    )

    `$fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    `$resolvedFontDir = `$null

    `$searchRoot = if (-not [string]::IsNullOrWhiteSpace(`$PSScriptRoot)) {
        `$PSScriptRoot
    } elseif (-not [string]::IsNullOrWhiteSpace(`$PSCommandPath)) {
        Split-Path -Path `$PSCommandPath -Parent
    } else {
        `$null
    }

    if (`$searchRoot) {
        `$probe = `$searchRoot
        for (`$i = 0; `$i -lt 5 -and -not [string]::IsNullOrWhiteSpace(`$probe); `$i++) {
            `$candidateDir = Join-Path -Path `$probe -ChildPath 'assets\fonts'
            if (Test-Path -Path (Join-Path -Path `$candidateDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                `$resolvedFontDir = `$candidateDir
                break
            }
            `$probe = Split-Path -Path `$probe -Parent
        }
    }

    if (-not `$resolvedFontDir) {
        `$cacheDir = Join-Path -Path `$env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        try {
            if (-not (Test-Path -Path `$cacheDir)) {
                New-Item -Path `$cacheDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            `$allCached = `$true
            foreach (`$ff in `$fontFiles) {
                if (-not (Test-Path -Path (Join-Path -Path `$cacheDir -ChildPath `$ff) -PathType Leaf)) {
                    `$allCached = `$false
                    break
                }
            }
            if (-not `$allCached) {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                `$prevProgress = `$ProgressPreference
                `$ProgressPreference = 'SilentlyContinue'
                try {
                    foreach (`$ff in `$fontFiles) {
                        `$dest = Join-Path -Path `$cacheDir -ChildPath `$ff
                        if (-not (Test-Path -Path `$dest -PathType Leaf)) {
                            `$url = "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/`$ff"
                            Invoke-WebRequest -Uri `$url -OutFile `$dest -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
                        }
                    }
                } finally {
                    `$ProgressPreference = `$prevProgress
                }
            }
            if (Test-Path -Path (Join-Path -Path `$cacheDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                `$resolvedFontDir = `$cacheDir
            }
        } catch {
            Write-Verbose "Figtree font download skipped: `$(`$_.Exception.Message)"
        }
    }

    if (`$resolvedFontDir) {
        `$fontUri = ([System.Uri]((`$resolvedFontDir.TrimEnd('\') + '\'))).AbsoluteUri
        `$Window.FontFamily = New-Object System.Windows.Media.FontFamily ("`$(`$fontUri)#Figtree, Segoe UI Variable, Segoe UI")
    } else {
        `$Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    if (-not `$PreserveSize) {
        `$workArea = [System.Windows.SystemParameters]::WorkArea
        `$targetWidth  = [Math]::Round([Math]::Min(`$workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, `$workArea.Width * 0.70))))
        `$targetHeight = [Math]::Round([Math]::Min(`$workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  `$workArea.Height * 0.80))))

        `$Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        `$Window.Width     = `$targetWidth
        `$Window.Height    = `$targetHeight
        `$Window.MinWidth  = [Math]::Min(860, [int]`$targetWidth)
        `$Window.MinHeight = [Math]::Min(640, [int]`$targetHeight)
        `$Window.Left      = [Math]::Round(`$workArea.Left + ((`$workArea.Width - `$targetWidth) / 2))
        `$Window.Top       = [Math]::Round(`$workArea.Top + ((`$workArea.Height - `$targetHeight) / 2))
        `$Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
    }
}
"@
} else { "" }

$bodyBlock = if ($Gui) {
@"

`$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$ToolName v1.0.0"
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

        <Style x:Key="PrimaryCtaStyle" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,0"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushAccentHover}"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushAccentPressed}"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryBtnStyle" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                                <Setter TargetName="BtnBorder" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
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

        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="$ToolName" FontSize="17" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextPrimary}"/>
                    <Border Margin="10,0,0,0" Background="{DynamicResource BrushAccentTint}" BorderBrush="{DynamicResource BrushAccentBorder}" BorderThickness="1" CornerRadius="4" Padding="8,2" VerticalAlignment="Center">
                        <TextBlock Text="v1.0.0" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                    </Border>
                </StackPanel>
                <Button x:Name="BtnThemeToggle" Grid.Column="1" Style="{StaticResource SecondaryBtnStyle}" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                    <TextBlock x:Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                </Button>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="24,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="$Synopsis" Foreground="{DynamicResource BrushTextSecondary}" FontSize="13" TextWrapping="Wrap" Margin="0,0,0,14"/>
            <Border Grid.Row="1" Background="{DynamicResource BrushLogBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" CornerRadius="6" Padding="12">
                <TextBox x:Name="TxtLogConsole"
                         Background="Transparent"
                         Foreground="{DynamicResource BrushTextSecondary}"
                         BorderThickness="0"
                         FontFamily="Cascadia Mono, Consolas"
                         FontSize="12"
                         IsReadOnly="True"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         Text="Ready."/>
            </Border>
        </Grid>

        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" VerticalAlignment="Center" Foreground="{DynamicResource BrushTextMuted}" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,0,16,0" Text="Logs: C:\Logs\$ToolName\$ToolName.log"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnClose" Style="{StaticResource SecondaryBtnStyle}" Content="Close" Margin="0,0,10,0"/>
                    <Button x:Name="BtnRun" Style="{StaticResource PrimaryCtaStyle}" Content="Run"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

try {
    Write-ToolLog -Message "Starting $ToolName GUI..." -Level INFO
    `$xamlReader = New-Object System.Xml.XmlNodeReader ([xml]`$xamlText)
    `$window = [Windows.Markup.XamlReader]::Load(`$xamlReader)
    Initialize-StandardGuiWindow -Window `$window

    `$txtThemeLabel = `$window.FindName('TxtThemeLabel')
    `$btnThemeToggle = `$window.FindName('BtnThemeToggle')
    `$txtLogConsole = `$window.FindName('TxtLogConsole')
    `$txtFooterHint = `$window.FindName('TxtFooterHint')
    `$btnClose = `$window.FindName('BtnClose')
    `$btnRun = `$window.FindName('BtnRun')

    `$Script:IsDarkTheme = Get-SystemIsDarkTheme
    `$Script:IsSystemThemeAuto = `$true
    Set-StandardGuiTheme -Window `$window -IsDark `$Script:IsDarkTheme -ThemeLabel `$txtThemeLabel -IsSystemAuto `$true

    `$window.Add_SourceInitialized({
        Set-StandardGuiTheme -Window `$window -IsDark `$Script:IsDarkTheme -ThemeLabel `$txtThemeLabel -IsSystemAuto `$Script:IsSystemThemeAuto
    })

    `$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
        param(`$sender, `$e)
        if (`$e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and `$Script:IsSystemThemeAuto) {
            `$window.Dispatcher.Invoke([Action]{
                `$Script:IsDarkTheme = Get-SystemIsDarkTheme
                Set-StandardGuiTheme -Window `$window -IsDark `$Script:IsDarkTheme -ThemeLabel `$txtThemeLabel -IsSystemAuto `$true
            })
        }
    }
    [Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged(`$userPrefHandler)
    `$window.Add_Closed({
        try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged(`$userPrefHandler) } catch {}
    })

    `$btnThemeToggle.Add_Click({
        `$Script:IsSystemThemeAuto = `$false
        `$Script:IsDarkTheme = -not `$Script:IsDarkTheme
        Set-StandardGuiTheme -Window `$window -IsDark `$Script:IsDarkTheme -ThemeLabel `$txtThemeLabel -IsSystemAuto `$false
    })

    `$btnClose.Add_Click({ `$window.Close() })
    `$btnRun.Add_Click({
        Write-ToolLog -Message "Running $ToolName..." -Level INFO
        `$txtLogConsole.AppendText("`r`n[`$(Get-Date -Format 'dd/MM/yy HH:mm:ss')] Executed $ToolName.")
        `$txtLogConsole.ScrollToEnd()
    })

    [void]`$window.ShowDialog()
    Write-ToolLog -Message "$ToolName completed successfully." -Level SUCCESS
} catch {
    Write-ToolLog -Message "Unhandled error: `$(`$_.Exception.Message)" -Level ERROR
    throw
}
"@
} else {
@"

try {
    Write-ToolLog -Message "Starting $ToolName..." -Level INFO
    # TODO: Implement core tool logic here
    Write-ToolLog -Message "$ToolName completed successfully." -Level SUCCESS
} catch {
    Write-ToolLog -Message "Unhandled error: `$(`$_.Exception.Message)" -Level ERROR
    throw
}
"@
}

if (-not (Test-Path -Path $scriptPath)) {
    $psContent = @"
<#
.SYNOPSIS
    $Synopsis

.DESCRIPTION
    $Synopsis
    Writes diagnostic execution logs to C:\Logs\$ToolName (with fallback to %LOCALAPPDATA%\Logs\$ToolName).

.EXAMPLE
    .\$ToolName.ps1

.NOTES
    Script Name   : $ToolName.ps1
    Author        : JP
    Version       : 1.0.0
    Created       : $shortDateStr
    Last Modified : $shortDateStr
    Log Path      : C:\Logs\$ToolName\$ToolName.log
#>
[CmdletBinding(SupportsShouldProcess = `$true)]
param()
$staBlock
`$ErrorActionPreference = 'Stop'

function Write-ToolLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = `$true)][string]`$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]`$Level = 'INFO'
    )
    `$logDir = "C:\Logs\$ToolName"
    if (-not (Test-Path -Path `$logDir)) {
        try {
            New-Item -Path `$logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            `$logDir = Join-Path `$env:LOCALAPPDATA "Logs\$ToolName"
            if (-not (Test-Path -Path `$logDir)) {
                New-Item -Path `$logDir -ItemType Directory -Force | Out-Null
            }
        }
    }
    `$timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    `$entry = "[`$timestamp] [`$Level] `$Message"
    Add-Content -Path (Join-Path `$logDir "$ToolName.log") -Value `$entry -Encoding UTF8
    Write-Verbose `$entry
}
$guiHelperBlock
$bodyBlock
"@
    [System.IO.File]::WriteAllText($scriptPath, $psContent, (New-Object System.Text.UTF8Encoding($false)))
}

if ($Gui) {
    $batPath = Join-Path $targetDir "Start-Gui.bat"
    if (-not (Test-Path -Path $batPath)) {
        $batContent = @"
@echo off
title $ToolName
echo ============================================================
echo Launching $ToolName GUI...
echo ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0$ToolName.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script closed with exit code %ERRORLEVEL%. Press any key to exit.
    pause
)
"@
        Set-Content -Path $batPath -Value $batContent -Encoding ASCII
    }
}

if (-not (Test-Path -Path $gitignorePath)) {
    [System.IO.File]::WriteAllText($gitignorePath, "Output/`r`n*.intunewin`r`n", (New-Object System.Text.UTF8Encoding($false)))
}

if (-not (Test-Path -Path $readmePath)) {
    $fastExecBlock = if ($Gui -or $Intune) {
@"
> [!WARNING]
> **Bundled Package Notice:**
> This tool is a multi-file package and cannot be executed in-memory via ``irm | iex``. Please download or clone the repository locally before executing.
>
> [!TIP]
> **Run locally in PowerShell (as Administrator):**
> Double-click ``Start-Gui.bat`` or run:
> ````powershell
> powershell.exe -ExecutionPolicy Bypass -File .\$ToolName.ps1
> ````
"@
    } else {
@"
> [!TIP]
> **Full Transparent GitHub Command:**
> ````powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/$Category/$ToolName/$ToolName.ps1")
> ````
"@
    }

    $readmeContent = @"
# $ToolName Guide

$Synopsis

## Overview

$Synopsis

### Key Features

- Standardized UK-timestamped diagnostic logging to ``C:\Logs\$ToolName\``.
- Structured error handling and PSScriptAnalyzer-compliant syntax.

## Prerequisites

- **OS Support:** Windows 10 (1909+) / Windows 11
- **PowerShell:** Windows PowerShell 5.1 or PowerShell 7+
- **Permissions:** Local Administrator privileges recommended

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Open an elevated PowerShell prompt or launch via ``Start-Gui.bat``.
2. Execute ``.\$ToolName.ps1``.
3. Review the execution output and log file in ``C:\Logs\$ToolName\$ToolName.log``.

### Logging & Diagnostics

| Log Path | Purpose |
| :--- | :--- |
| ``C:\Logs\$ToolName\$ToolName.log`` | Primary execution and error audit log |

## Fast Execute

$fastExecBlock
"@
    [System.IO.File]::WriteAllText($readmePath, $readmeContent, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Output "Scaffolded wrapped PowerShell tool at: $targetDir"
