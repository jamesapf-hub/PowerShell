<#
.SYNOPSIS
    Main runner script for the Ultimate PC CleanUp Utility.
.DESCRIPTION
    This script can be run in two modes:
    1. Interactive GUI mode (default): Launches a modern dark-themed WPF dashboard.
    2. Command-Line mode (via switches): Performs cleanup directly in the terminal,
       making it suitable for RMM, scheduled tasks, or shell automation.
.PARAMETER NoGui
    Launches the tool in command-line mode without displaying the GUI window.
.PARAMETER RunAll
    Automatically runs all discovered tasks in command-line mode.
.PARAMETER RunTasks
    An array of task basenames/filenames to execute in command-line mode.
.PARAMETER ListTasks
    Lists all available tasks with descriptions in the terminal and exits.
.NOTES
    Author  : Antigravity Pair Programmer
    Version : 1.1
#>

[CmdletBinding()]
param (
    [switch]$NoGui,
    [switch]$RunAll,
    [string[]]$RunTasks,
    [switch]$ListTasks,
    [switch]$Gui
)

# Set base parameters
try {
    if (-not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Definition)) {
        $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition -ErrorAction Stop
    } else {
        $PSScriptRoot = ""
    }
} catch {
    $PSScriptRoot = ""
}

# CRITICAL SAFETY GUARD: Prevent script from running in memory via iex
if ([string]::IsNullOrEmpty($PSScriptRoot) -or $PSScriptRoot -match "^iex") {
    Write-Host ""
    Write-Host "==========================================================" -ForegroundColor Red
    Write-Host "CRITICAL ERROR: UltimateCleanUp Cannot Be Run From Memory!" -ForegroundColor Red
    Write-Host "==========================================================" -ForegroundColor Red
    Write-Host "UltimateCleanUp is a packaged utility that requires local" -ForegroundColor Yellow
    Write-Host "copies of all sub-scripts to build its dashboard." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please use the 'Package & Download' feature in the portal" -ForegroundColor White
    Write-Host "to download the entire UltimateCleanUp folder, extract it," -ForegroundColor White
    Write-Host "and double-click Run-UltimateCleanUp.bat locally." -ForegroundColor White
    Write-Host "==========================================================" -ForegroundColor Red
    exit 1
}

$ScriptsFolder = Join-Path -Path $PSScriptRoot -ChildPath "Scripts"

# Ensure Scripts folder exists
if (-not (Test-Path $ScriptsFolder)) {
    New-Item -ItemType Directory -Path $ScriptsFolder | Out-Null
}

# Default behavior: If no parameters are specified, default to ListTasks (CLI mode)
$noParamsSpecified = -not ($PSBoundParameters.ContainsKey('NoGui') -or 
                           $PSBoundParameters.ContainsKey('RunAll') -or 
                           $PSBoundParameters.ContainsKey('RunTasks') -or 
                           $PSBoundParameters.ContainsKey('ListTasks') -or 
                           $PSBoundParameters.ContainsKey('Gui'))

if ($noParamsSpecified) {
    $Gui = $true
    $NoGui = $false
}

# If any command-line execution parameter is passed, default to NoGui
if ($RunAll -or $RunTasks -or $ListTasks) {
    $NoGui = $true
}

# If GUI is explicitly requested, override NoGui
if ($Gui) {
    $NoGui = $false
}

# Check if running as administrator (required for system-level cleanup tasks)
$currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Request UAC Elevation
    if ($PSCommandPath) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex (irm 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/UltimateCleanUp/Start-UltimateCleanUp.ps1')`""
    }
    if ($NoGui) { $arguments += " -NoGui" }
    if ($RunAll) { $arguments += " -RunAll" }
    if ($ListTasks) { $arguments += " -ListTasks" }
    if ($Gui) { $arguments += " -Gui" }
    if ($RunTasks) {
        $arguments += " -RunTasks $($RunTasks -join ',')"
    }
    
    try {
        $currentDir = (Get-Location).Path
        if ($NoGui) {
            # In NoGui/CLI mode, we want to run in the current console and wait for exit
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -WorkingDirectory $currentDir -Wait
        } else {
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -WorkingDirectory $currentDir
        }
    } catch {
        Write-Error "This utility requires Administrator privileges."
    }
    exit
}

# Reusable function to discover scripts and extract synopsis metadata
function Get-DiscoveredScript {
    if (-not (Test-Path $ScriptsFolder)) {
        return @()
    }
    
    $files = Get-ChildItem -Path $ScriptsFolder -Filter *.ps1 -Recurse | Sort-Object Name
    $list = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($file in $files) {
        # Parse synopsis block if present
        $content = Get-Content $file.FullName -Head 20
        $synopsis = ""
        $inSynopsis = $false
        foreach ($line in $content) {
            if ($line -match '\.SYNOPSIS') {
                $inSynopsis = $true
                continue
            }
            if ($inSynopsis) {
                if ($line -match '^\s*\.\w+' -or $line -match '#>' -or $line -match '^\s*$') {
                    $inSynopsis = $false
                } else {
                    $synopsis += ($line.Trim('#').Trim() + " ")
                }
            }
        }
        
        $synopsis = $synopsis.Trim()
        if (-not $synopsis) {
            $synopsis = "No script synopsis details provided."
        }
        
        # Format friendly title (e.g. "01_ClearTempFiles.ps1" -> "Clear Temp Files")
        $title = $file.BaseName -replace '^\d+[-_]', '' -replace '[-_]', ' '
        $title = $title -creplace '([a-z])([A-Z])', '$1 $2'
        $title = (Get-Culture).TextInfo.ToTitleCase($title)
        
        $list.Add([PSCustomObject]@{
            Path        = $file.FullName
            FileName    = $file.Name
            BaseName    = $file.BaseName
            Title       = $title
            Description = $synopsis
        }) | Out-Null
    }
    return $list
}

# ==========================================
# COMMAND LINE / AUTOMATION MODE
# ==========================================
if ($NoGui) {
    $tasks = Get-DiscoveredScript
    if ($tasks.Count -eq 0) {
        Write-Host "[-] No cleanup tasks found in '$ScriptsFolder'." -ForegroundColor Red
        exit 0
    }
    
    if ($ListTasks) {
        Write-Host "`r`n=== DISCOVERED CLEANUP TASKS ===" -ForegroundColor Cyan
        foreach ($task in $tasks) {
            Write-Host "- [$($task.BaseName)] $($task.Title)" -ForegroundColor Yellow
            Write-Host "  Description: $($task.Description)`r`n" -ForegroundColor DarkGray
        }
        exit 0
    }
    
    # Filter tasks to run
    $tasksToRun = @()
    if ($RunTasks) {
        foreach ($target in $RunTasks) {
            $match = $tasks | Where-Object { $_.BaseName -eq $target -or $_.Title -eq $target -or $_.FileName -eq $target }
            if ($match) {
                $tasksToRun += $match
            } else {
                Write-Host "[-] Warning: Task '$target' not found." -ForegroundColor Yellow
            }
        }
    } else {
        # Default to running all if RunAll or no parameter specified
        $tasksToRun = $tasks
    }
    
    if ($tasksToRun.Count -eq 0) {
        Write-Host "[-] No valid tasks to execute." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`r`n==================================================" -ForegroundColor Cyan
    Write-Host "STARTING COMMAND LINE PC CLEANUP" -ForegroundColor Cyan
    Write-Host "Executing $($tasksToRun.Count) tasks..." -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    
    $successCount = 0
    foreach ($task in $tasksToRun) {
        Write-Host "`r`n>>> Running task: $($task.Title)..." -ForegroundColor Yellow
        
        # Start background powershell instance to run the task
        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($task.Path)`" -Force" -NoNewWindow -Wait -PassThru
        
        if ($proc.ExitCode -eq 0) {
            Write-Host "[+] Task '$($task.Title)' completed successfully." -ForegroundColor Green
            $successCount++
        } else {
            Write-Host "[-] Task '$($task.Title)' failed with exit code $($proc.ExitCode)." -ForegroundColor Red
        }
    }
    
    Write-Host "`r`n==================================================" -ForegroundColor Cyan
    Write-Host "CLEANUP FINISHED: $successCount of $($tasksToRun.Count) tasks completed successfully." -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    exit 0
}

# ==========================================
# INTERACTIVE GUI MODE (WPF)
# ==========================================

# Enforce STA ApartmentState for WPF UI rendering (only if GUI mode is active)
if (-not $NoGui -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Verbose "Thread is not STA. Relaunching GUI in a new STA process..."
    if ($PSCommandPath) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -File `"$PSCommandPath`""
    } else {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex (irm 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/UltimateCleanUp/Start-UltimateCleanUp.ps1')`""
    }
    if ($Gui) { $arguments += " -Gui" }
    
    $currentDir = (Get-Location).Path
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory $currentDir -Wait
    return
}

# Import GUI assemblies
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
            BrushChromeBar            = '#CC070D19'
            BrushHairline             = '#1CFFFFFF'
            BrushCardBg               = '#B30D192E'
            BrushCardHoverBg          = '#CC13223A'
            BrushCardBorder           = '#26354D'
            BrushInputBg              = '#C7070D1A'
            BrushLogBg                = '#E6050A14'
            BrushConsoleHeaderBg      = '#660D192E'
            BrushTextPrimary          = '#F8FAFC'
            BrushTextSecondary        = '#CBD5E1'
            BrushTextMuted            = '#94A3B8'
            BrushTextDisabled         = '#64748B'
            BrushAccent               = '#2DD4BF'
            BrushAccentHover          = '#5EEAD4'
            BrushAccentPressed        = '#14B8A6'
            BrushAccentFg             = '#042F2E'
            BrushAccentTint           = '#242DD4BF'
            BrushAccentBorder         = '#6B2DD4BF'
            BrushSecBtnBg             = '#14FFFFFF'
            BrushSecBtnHover          = '#24FFFFFF'
            BrushSecBtnPressed        = '#1FFFFFFF'
            BrushBtnDisabledBg        = '#13223A'
            BrushSecBtnDisabledBg     = '#0DFFFFFF'
            BrushSecBtnDisabledBorder = '#1A26354D'
            BrushScrollThumb          = '#26354D'
            BrushProgressTrack        = '#99070D1A'
            BrushBadgePendingBg       = '#99070D1A'
            BrushBadgePendingFg       = '#94A3B8'
            BrushBadgeRunningBg       = '#1A2DD4BF'
            BrushBadgeRunningFg       = '#2DD4BF'
            BrushBadgeSuccessBg       = '#064E3B'
            BrushBadgeSuccessFg       = '#34D399'
            BrushBadgeFailedBg        = '#7F1D1D'
            BrushBadgeFailedFg        = '#F87171'
        }
    } else {
        $canvasBrush = $conv.ConvertFromString('#F8F9FF')
        $map = @{
            BrushChromeBar            = '#F2FFFFFF'
            BrushHairline             = '#190F172A'
            BrushCardBg               = '#FFFFFFFF'
            BrushCardHoverBg          = '#FFF1F5F9'
            BrushCardBorder           = '#CBD5E1'
            BrushInputBg              = '#F8FAFC'
            BrushLogBg                = '#F1F5F9'
            BrushConsoleHeaderBg      = '#F8FAFC'
            BrushTextPrimary          = '#0F172A'
            BrushTextSecondary        = '#334155'
            BrushTextMuted            = '#64748B'
            BrushTextDisabled         = '#94A3B8'
            BrushAccent               = '#2563EB'
            BrushAccentHover          = '#1D4ED8'
            BrushAccentPressed        = '#1E40AF'
            BrushAccentFg             = '#FFFFFF'
            BrushAccentTint           = '#DBEAFE'
            BrushAccentBorder         = '#592563EB'
            BrushSecBtnBg             = '#FFFFFF'
            BrushSecBtnHover          = '#F1F5F9'
            BrushSecBtnPressed        = '#E2E8F0'
            BrushBtnDisabledBg        = '#E2E8F0'
            BrushSecBtnDisabledBg     = '#F1F5F9'
            BrushSecBtnDisabledBorder = '#CBD5E1'
            BrushScrollThumb          = '#CBD5E1'
            BrushProgressTrack        = '#E2E8F0'
            BrushBadgePendingBg       = '#F1F5F9'
            BrushBadgePendingFg       = '#475569'
            BrushBadgeRunningBg       = '#DBEAFE'
            BrushBadgeRunningFg       = '#1D4ED8'
            BrushBadgeSuccessBg       = '#DCFCE7'
            BrushBadgeSuccessFg       = '#15803D'
            BrushBadgeFailedBg        = '#FEE2E2'
            BrushBadgeFailedFg        = '#B91C1C'
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
        [string]$ScriptRoot = $PSScriptRoot
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $fontBaseUrl = 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/'
    $resolvedFontDir = $null

    if (-not [string]::IsNullOrWhiteSpace($ScriptRoot)) {
        $probeDir = $ScriptRoot
        for ($level = 0; $level -lt 6; $level++) {
            if ([string]::IsNullOrWhiteSpace($probeDir)) { break }
            $candidateDir = Join-Path -Path $probeDir -ChildPath 'assets\fonts'
            if (Test-Path -LiteralPath (Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf')) {
                $resolvedFontDir = $candidateDir
                break
            }
            $probeDir = Split-Path -Path $probeDir -Parent
        }
    }

    if (-not $resolvedFontDir) {
        $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        try {
            if (-not (Test-Path -LiteralPath $cacheDir)) {
                New-Item -Path $cacheDir -ItemType Directory -Force | Out-Null
            }
            $prevProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
                foreach ($fontFile in $fontFiles) {
                    $destPath = Join-Path -Path $cacheDir -ChildPath $fontFile
                    if (-not (Test-Path -LiteralPath $destPath)) {
                        Invoke-WebRequest -Uri ($fontBaseUrl + $fontFile) -OutFile $destPath -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                    }
                }
            } finally {
                $ProgressPreference = $prevProgress
            }
            if (Test-Path -LiteralPath (Join-Path -Path $cacheDir -ChildPath 'Figtree-Regular.ttf')) {
                $resolvedFontDir = $cacheDir
            }
        } catch {
            $resolvedFontDir = $null
        }
    }

    if ($resolvedFontDir -and (Test-Path -LiteralPath (Join-Path -Path $resolvedFontDir -ChildPath 'Figtree-Regular.ttf'))) {
        try {
            $fontUri = ([System.Uri]($resolvedFontDir.TrimEnd('\') + '\')).AbsoluteUri
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
        } catch {
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("Segoe UI Variable, Segoe UI")
        }
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("Segoe UI Variable, Segoe UI")
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Window.Width = $targetWidth
    $Window.Height = $targetHeight
    $Window.MinWidth = [Math]::Min(860, $targetWidth)
    $Window.MinHeight = [Math]::Min(640, $targetHeight)
    $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
    $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
    $Window.Left = $workArea.Left + [Math]::Round(($workArea.Width - $targetWidth) / 2)
    $Window.Top  = $workArea.Top  + [Math]::Round(($workArea.Height - $targetHeight) / 2)
}

# Define XAML for the Main Window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ultimate PC CleanUp Utility v1.1"
        Height="720" Width="1120"
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
        <SolidColorBrush x:Key="BrushCardHoverBg" Color="#CC13223A"/>
        <SolidColorBrush x:Key="BrushCardBorder" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushInputBg" Color="#C7070D1A"/>
        <SolidColorBrush x:Key="BrushLogBg" Color="#E6050A14"/>
        <SolidColorBrush x:Key="BrushConsoleHeaderBg" Color="#660D192E"/>
        <SolidColorBrush x:Key="BrushTextPrimary" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="BrushTextSecondary" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="BrushTextMuted" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BrushTextDisabled" Color="#64748B"/>
        <SolidColorBrush x:Key="BrushAccent" Color="#2DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentHover" Color="#5EEAD4"/>
        <SolidColorBrush x:Key="BrushAccentPressed" Color="#14B8A6"/>
        <SolidColorBrush x:Key="BrushAccentFg" Color="#042F2E"/>
        <SolidColorBrush x:Key="BrushAccentTint" Color="#242DD4BF"/>
        <SolidColorBrush x:Key="BrushAccentBorder" Color="#6B2DD4BF"/>
        <SolidColorBrush x:Key="BrushSecBtnBg" Color="#14FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnHover" Color="#24FFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnPressed" Color="#1FFFFFFF"/>
        <SolidColorBrush x:Key="BrushBtnDisabledBg" Color="#13223A"/>
        <SolidColorBrush x:Key="BrushSecBtnDisabledBg" Color="#0DFFFFFF"/>
        <SolidColorBrush x:Key="BrushSecBtnDisabledBorder" Color="#1A26354D"/>
        <SolidColorBrush x:Key="BrushScrollThumb" Color="#26354D"/>
        <SolidColorBrush x:Key="BrushProgressTrack" Color="#99070D1A"/>
        <SolidColorBrush x:Key="BrushBadgePendingBg" Color="#99070D1A"/>
        <SolidColorBrush x:Key="BrushBadgePendingFg" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BrushBadgeRunningBg" Color="#1A2DD4BF"/>
        <SolidColorBrush x:Key="BrushBadgeRunningFg" Color="#2DD4BF"/>
        <SolidColorBrush x:Key="BrushBadgeSuccessBg" Color="#064E3B"/>
        <SolidColorBrush x:Key="BrushBadgeSuccessFg" Color="#34D399"/>
        <SolidColorBrush x:Key="BrushBadgeFailedBg" Color="#7F1D1D"/>
        <SolidColorBrush x:Key="BrushBadgeFailedFg" Color="#F87171"/>

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

        <!-- Primary Action Button Style -->
        <Style TargetType="Button" x:Key="PrimaryBtn">
            <Setter Property="Background" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccentFg}"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushBtnDisabledBg}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextDisabled}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Accent Outline Button Style -->
        <Style TargetType="Button" x:Key="AccentBtn">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushSecBtnHover}"/>
                                <Setter TargetName="BtnBorder" Property="BorderBrush" Value="{DynamicResource BrushAccentHover}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushAccentTint}"/>
                                <Setter TargetName="BtnBorder" Property="BorderBrush" Value="{DynamicResource BrushAccentPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushSecBtnDisabledBg}"/>
                                <Setter TargetName="BtnBorder" Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextDisabled}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- General Secondary Button Style -->
        <Style TargetType="Button" x:Key="DefaultBtn">
            <Setter Property="Background" Value="{DynamicResource BrushSecBtnBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource BrushTextSecondary}"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="11.5"/>
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
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushSecBtnPressed}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="BtnBorder" Property="Background" Value="{DynamicResource BrushSecBtnDisabledBg}"/>
                                <Setter TargetName="BtnBorder" Property="BorderBrush" Value="{DynamicResource BrushSecBtnDisabledBorder}"/>
                                <Setter Property="Foreground" Value="{DynamicResource BrushTextDisabled}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dynamic CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource BrushTextPrimary}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Grid Background="Transparent">
                            <Border x:Name="CheckBoxBorder" Width="18" Height="18" Background="{DynamicResource BrushInputBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1.5" CornerRadius="4" VerticalAlignment="Center" HorizontalAlignment="Center">
                                <Path x:Name="CheckMark" Data="M 3,8.5 L 6.5,12 L 13,4.5" Stroke="{DynamicResource BrushAccentFg}" StrokeThickness="2.2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckBoxBorder" Property="Background" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="CheckBoxBorder" Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckBoxBorder" Property="BorderBrush" Value="{DynamicResource BrushAccentHover}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Script Card Item Style -->
        <Style TargetType="Border" x:Key="ScriptCardStyle">
            <Setter Property="Background" Value="{DynamicResource BrushCardBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BrushCardBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="Padding" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource BrushCardHoverBg}"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource BrushAccent}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <!-- Header Bar -->
            <RowDefinition Height="Auto"/>
            <!-- Main Content Area -->
            <RowDefinition Height="*"/>
            <!-- Footer Bar -->
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Panel -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderThickness="0,0,0,1" BorderBrush="{DynamicResource BrushHairline}" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Ultimate PC CleanUp" FontSize="20" FontWeight="Bold" Foreground="{DynamicResource BrushTextPrimary}" VerticalAlignment="Center"/>
                        <Border Background="{DynamicResource BrushAccentTint}" BorderBrush="{DynamicResource BrushAccentBorder}" BorderThickness="1" CornerRadius="4" Padding="7,2" Margin="10,0,0,0" VerticalAlignment="Center">
                            <TextBlock Text="v1.1" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas"/>
                        </Border>
                    </StackPanel>
                    <TextBlock Text="Run custom PC tidy-up scripts safely, monitor outputs, and keep Windows clean." FontSize="12.5" Foreground="{DynamicResource BrushTextMuted}" Margin="0,4,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="BtnThemeToggle" Style="{StaticResource DefaultBtn}" Height="34" Margin="0,0,10,0" Cursor="Hand" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                        <TextBlock Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                    </Button>
                    <Button Name="BtnRefresh" Style="{StaticResource DefaultBtn}" Content="Scan Folder" MinWidth="106" Height="34" Cursor="Hand" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main Layout Split -->
        <Grid Grid.Row="1" Margin="24,18">
            <Grid.ColumnDefinitions>
                <!-- Task Checklist Selection (Left) -->
                <ColumnDefinition Width="4*"/>
                <!-- Vertical Divider -->
                <ColumnDefinition Width="18"/>
                <!-- Console Logging Console (Right) -->
                <ColumnDefinition Width="5*"/>
            </Grid.ColumnDefinitions>

            <!-- Left Panel: Available Scripts -->
            <Grid Grid.Column="0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Text="DISCOVERED CLEANUP TASKS" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushTextMuted}" Margin="0,0,0,10"/>

                <!-- Scrollable Checklist -->
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                    <StackPanel Name="ScriptListContainer" Margin="0,0,6,0"/>
                </ScrollViewer>

                <!-- Bulk Selection controls -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
                    <Button Name="BtnSelectAll" Style="{StaticResource DefaultBtn}" Content="Select All" Margin="0,0,8,0" Padding="14,6" Cursor="Hand"/>
                    <Button Name="BtnDeselectAll" Style="{StaticResource DefaultBtn}" Content="Clear Selection" Padding="14,6" Cursor="Hand"/>
                </StackPanel>
            </Grid>

            <!-- Right Panel: Logs & Execution Controls -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Execution Buttons -->
                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Name="BtnRunSelected" Style="{StaticResource AccentBtn}" Content="Run Selected Tasks" Margin="0,0,10,0" Height="40" Cursor="Hand"/>
                    <Button Name="BtnRunAll" Grid.Column="1" Style="{StaticResource PrimaryBtn}" Content="Run All Tasks" Height="40" Cursor="Hand"/>
                </Grid>

                <!-- Monospaced Log Console -->
                <Border Grid.Row="1" Background="{DynamicResource BrushLogBg}" CornerRadius="6" BorderThickness="1" BorderBrush="{DynamicResource BrushCardBorder}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Console Header -->
                        <Border Background="{DynamicResource BrushConsoleHeaderBg}" Padding="14,8" BorderThickness="0,0,0,1" BorderBrush="{DynamicResource BrushCardBorder}">
                            <Grid>
                                <TextBlock Text="REAL-TIME EXECUTION LOG" FontSize="11" FontWeight="Bold" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center"/>
                                <Button Name="BtnClearLog" Content="Clear Console" Style="{StaticResource DefaultBtn}" HorizontalAlignment="Right" Padding="10,4" FontSize="11" Cursor="Hand"/>
                            </Grid>
                        </Border>

                        <!-- Console Textbox -->
                        <TextBox Name="TxtConsole" Grid.Row="1" Background="Transparent" Foreground="{DynamicResource BrushTextPrimary}" CaretBrush="{DynamicResource BrushAccent}" SelectionBrush="{DynamicResource BrushAccent}" FontFamily="Cascadia Mono, Consolas" FontSize="12.5" BorderThickness="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" Margin="12,10" Padding="2"/>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- Footer Panel -->
        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderThickness="0,1,0,0" BorderBrush="{DynamicResource BrushHairline}" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="TxtProgress" Text="Ready to begin cleanup" Foreground="{DynamicResource BrushTextMuted}" VerticalAlignment="Center" FontSize="12"/>
                <ProgressBar Name="ProgressIndicator" Grid.Column="1" Margin="20,0" Height="8" Background="{DynamicResource BrushProgressTrack}" Foreground="{DynamicResource BrushAccent}" Minimum="0" Maximum="100" Value="0" BorderThickness="0">
                    <ProgressBar.Template>
                        <ControlTemplate TargetType="ProgressBar">
                            <Border Background="{TemplateBinding Background}" CornerRadius="4" ClipToBounds="True">
                                <Grid>
                                    <Border x:Name="PART_Track" />
                                    <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left" CornerRadius="4"/>
                                </Grid>
                            </Border>
                        </ControlTemplate>
                    </ProgressBar.Template>
                </ProgressBar>
                <TextBlock Name="TxtStatusCount" Grid.Column="2" Text="0 / 0 Completed" Foreground="{DynamicResource BrushTextSecondary}" VerticalAlignment="Center" FontSize="12" FontWeight="SemiBold" FontFamily="Cascadia Mono, Consolas"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Load the XML XAML into WPF Window object
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $window -ScriptRoot $PSScriptRoot

# Map XAML named controls to PowerShell script-scoped variables automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name $_.Name -Value $window.FindName($_.Name) -Scope Script
}

# Theme State + Painted Win11 Chrome Hook
$Script:IsDarkTheme = Get-SystemIsDarkTheme
$Script:IsSystemThemeAuto = $true
Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $true

$window.Add_SourceInitialized({
    Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $Script:IsSystemThemeAuto
})

$userPrefHandler = [Microsoft.Win32.UserPreferenceChangedEventHandler]{
    param($sender, $e)
    if ($e.Category -eq [Microsoft.Win32.UserPreferenceCategory]::General -and $Script:IsSystemThemeAuto) {
        $window.Dispatcher.Invoke([Action]{
            $Script:IsDarkTheme = Get-SystemIsDarkTheme
            Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $true
        })
    }
}
[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged($userPrefHandler)
$window.Add_Closed({
    try { [Microsoft.Win32.SystemEvents]::remove_UserPreferenceChanged($userPrefHandler) } catch {}
})

$BtnThemeToggle.Add_Click({
    $Script:IsSystemThemeAuto = $false
    $Script:IsDarkTheme = -not $Script:IsDarkTheme
    Set-StandardGuiTheme -Window $window -IsDark $Script:IsDarkTheme -ThemeLabel $TxtThemeLabel -IsSystemAuto $false
})

# Synchronized hashtable to share state with background execution thread safely
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Window = $window
$syncHash.TxtConsole = $TxtConsole
$syncHash.ProgressIndicator = $ProgressIndicator
$syncHash.TxtStatusCount = $TxtStatusCount
$syncHash.BtnRunAll = $BtnRunAll
$syncHash.BtnRunSelected = $BtnRunSelected
$syncHash.BtnRefresh = $BtnRefresh
$syncHash.TxtProgress = $TxtProgress

# Initialize a list to hold discovered script objects
$scriptList = [System.Collections.Generic.List[PSCustomObject]]::new()

# Create a background runspace pool / runspace to execute scripts asynchronously
$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

# Function to parse scripts inside the 'Scripts' folder
function Update-ScriptList {
    $ScriptListContainer.Children.Clear()
    $scriptList.Clear()
    
    $discovered = Get-DiscoveredScript
    foreach ($file in $discovered) {
        $safeTitle = [System.Security.SecurityElement]::Escape($file.Title)
        $safeDescription = [System.Security.SecurityElement]::Escape($file.Description)
        # Load XAML for the dynamic Card UI element
        $cardXaml = @"
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Style="{DynamicResource ScriptCardStyle}">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        
        <CheckBox Name="ChkSelect" VerticalAlignment="Center" Margin="0,0,12,0" IsChecked="True" Cursor="Hand"/>
        
        <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="0,0,10,0">
            <TextBlock Name="TxtTitle" Text="$safeTitle" FontWeight="SemiBold" FontSize="13.5" Foreground="{DynamicResource BrushTextPrimary}"/>
            <TextBlock Name="TxtDesc" Text="$safeDescription" FontSize="11.5" Foreground="{DynamicResource BrushTextMuted}" TextWrapping="Wrap" Margin="0,3,0,0"/>
        </StackPanel>
        
        <Border Grid.Column="2" Name="StatusBadge" CornerRadius="4" Padding="8,4" VerticalAlignment="Center" Background="{DynamicResource BrushBadgePendingBg}" BorderBrush="{DynamicResource BrushCardBorder}" BorderThickness="1" MinWidth="75">
            <TextBlock Name="TxtStatus" Text="Pending" Foreground="{DynamicResource BrushBadgePendingFg}" FontWeight="Bold" FontSize="10.5" HorizontalAlignment="Center"/>
        </Border>
    </Grid>
</Border>
"@
        $card = [Windows.Markup.XamlReader]::Parse($cardXaml)
        $ScriptListContainer.Children.Add($card)
        
        # Extract dynamic card elements
        $scriptObj = [PSCustomObject]@{
            Path        = $file.Path
            FileName    = $file.FileName
            Title       = $file.Title
            Description = $file.Description
            Card        = $card
            Checkbox    = $card.FindName("ChkSelect")
            StatusText  = $card.FindName("TxtStatus")
            StatusBadge = $card.FindName("StatusBadge")
        }
        $scriptList.Add($scriptObj) | Out-Null
    }
    
    # Update Footer statuses
    $TxtStatusCount.Text = "0 / $($scriptList.Count) Completed"
    $ProgressIndicator.Value = 0
}

# Function to kick off execution (Runs on background Runspace thread)
function Start-Execution {
    param([bool]$onlySelected)
    
    # Identify scripts to execute
    $selectedScripts = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($script in $scriptList) {
        # Reset card status visually
        $script.StatusText.Text = "Pending"
        $script.StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushBadgePendingBg')
        $script.StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushBadgePendingFg')
        
        if (-not $onlySelected -or $script.Checkbox.IsChecked) {
            $selectedScripts.Add($script) | Out-Null
        }
    }
    
    if ($selectedScripts.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one task to run.", "No Tasks Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Lock controls during execution
    $BtnRunAll.IsEnabled = $false
    $BtnRunSelected.IsEnabled = $false
    $BtnRefresh.IsEnabled = $false
    
    $TxtProgress.Text = "Executing cleanup..."
    $ProgressIndicator.Value = 0
    
    # Share state with background thread
    $syncHash.Scripts = $selectedScripts
    $syncHash.TotalCount = $selectedScripts.Count
    $syncHash.CompletedCount = 0
    
    # Reset console log
    $TxtConsole.Clear()
    $TxtConsole.AppendText("[*] Starting PC cleanup utility process...`r`n`r`n")
    
    # Setup PowerShell script execution engine
    $ps = [PowerShell]::Create()
    $ps.Runspace = $runspace
    
    $runscript = {
        $scripts = $syncHash.Scripts
        
        for ($i = 0; $i -lt $scripts.Count; $i++) {
            $script = $scripts[$i]
            
            # Transition task card UI status to 'Running'
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $script.StatusText.Text = "Running"
                $script.StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushBadgeRunningBg')
                $script.StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushBadgeRunningFg')
                
                $syncHash.TxtConsole.AppendText("======================================================================`r`n")
                $syncHash.TxtConsole.AppendText("TASK: $($script.Title)`r`n")
                $syncHash.TxtConsole.AppendText("======================================================================`r`n")
                $syncHash.TxtConsole.ScrollToEnd()
            })
            
            # Prepare process start properties (run sub-script inside hidden process)
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "powershell.exe"
            $pinfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($script.Path)`" -Force"
            $pinfo.RedirectStandardOutput = $true
            $pinfo.RedirectStandardError = $true
            $pinfo.UseShellExecute = $false
            $pinfo.CreateNoWindow = $true
            
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $pinfo
            
            $proc.Start() | Out-Null
            
            # Read stdout stream asynchronously line-by-line
            while (-not $proc.HasExited) {
                $line = $proc.StandardOutput.ReadLine()
                if ($null -ne $line) {
                    $syncHash.Window.Dispatcher.Invoke([Action]{
                        $syncHash.TxtConsole.AppendText("$line`r`n")
                        $syncHash.TxtConsole.ScrollToEnd()
                    })
                }
                Start-Sleep -Milliseconds 10
            }
            
            # Capture lingering standard outputs
            $remaining = $proc.StandardOutput.ReadToEnd()
            if (-not [string]::IsNullOrEmpty($remaining)) {
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    $syncHash.TxtConsole.AppendText("$remaining")
                    $syncHash.TxtConsole.ScrollToEnd()
                })
            }
            
            # Capture errors
            $errors = $proc.StandardError.ReadToEnd()
            $exitCode = $proc.ExitCode
            
            # Update Task card UI state based on exit code
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.CompletedCount++
                $percent = [math]::Round(($syncHash.CompletedCount / $syncHash.TotalCount) * 100)
                $syncHash.ProgressIndicator.Value = $percent
                $syncHash.TxtStatusCount.Text = "$($syncHash.CompletedCount) / $($syncHash.TotalCount) Completed"
                
                if ($exitCode -eq 0) {
                    $script.StatusText.Text = "Success"
                    $script.StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushBadgeSuccessBg')
                    $script.StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushBadgeSuccessFg')
                    $syncHash.TxtConsole.AppendText("`r`n[+] Task completed successfully.`r`n`r`n")
                } else {
                    $script.StatusText.Text = "Failed"
                    $script.StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BrushBadgeFailedBg')
                    $script.StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'BrushBadgeFailedFg')
                    if (-not [string]::IsNullOrEmpty($errors)) {
                        $syncHash.TxtConsole.AppendText("`r`n[!] Errors reported:`r`n$errors`r`n")
                    }
                    $syncHash.TxtConsole.AppendText("`r`n[-] Task failed with exit code $exitCode.`r`n`r`n")
                }
                $syncHash.TxtConsole.ScrollToEnd()
            })
        }
        
        # Enable controls after tasks finish
        $syncHash.Window.Dispatcher.Invoke([Action]{
            $syncHash.BtnRunAll.IsEnabled = $true
            $syncHash.BtnRunSelected.IsEnabled = $true
            $syncHash.BtnRefresh.IsEnabled = $true
            $syncHash.TxtProgress.Text = "All tasks processed."
            $syncHash.TxtConsole.AppendText("======================================================================`r`n")
            $syncHash.TxtConsole.AppendText("CLEANUP PROCESS FINISHED`r`n")
            $syncHash.TxtConsole.AppendText("======================================================================`r`n")
            $syncHash.TxtConsole.ScrollToEnd()
        })
    }
    
    $ps.AddScript($runscript).BeginInvoke() | Out-Null
}

# --- Event Handlers ---

# Refresh/Scan Folder Button
$BtnRefresh.Add_Click({
    Update-ScriptList
})

# Run Selected Tasks Button
$BtnRunSelected.Add_Click({
    Start-Execution -onlySelected $true
})

# Run All Tasks Button
$BtnRunAll.Add_Click({
    Start-Execution -onlySelected $false
})

# Clear Log Button
$BtnClearLog.Add_Click({
    $TxtConsole.Clear()
})

# Select All Tasks
$BtnSelectAll.Add_Click({
    foreach ($script in $scriptList) {
        $script.Checkbox.IsChecked = $true
    }
})

# Clear/Deselect All Tasks
$BtnDeselectAll.Add_Click({
    foreach ($script in $scriptList) {
        $script.Checkbox.IsChecked = $false
    }
})

# Window closing event - clean up runspaces
$window.Add_Closing({
    try {
        if ($runspace) {
            $runspace.Close()
            $runspace.Dispose()
        }
    } catch {}
})

# Initial script scan and load
Update-ScriptList

# Display the main window
$window.ShowDialog() | Out-Null
