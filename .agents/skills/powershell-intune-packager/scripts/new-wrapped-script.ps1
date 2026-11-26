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
        Foreground="#F8FAFC"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True">
    <Window.Background>
        <RadialGradientBrush Center="0.5,0.0" GradientOrigin="0.5,0.0" RadiusX="0.95" RadiusY="0.85">
            <GradientStop Color="#0F1F3D" Offset="0.0"/>
            <GradientStop Color="#091122" Offset="0.55"/>
            <GradientStop Color="#05070C" Offset="1.0"/>
        </RadialGradientBrush>
    </Window.Background>
    <Window.Resources>
        <Style x:Key="PrimaryCtaStyle" TargetType="Button">
            <Setter Property="Background" Value="#2DD4BF"/>
            <Setter Property="Foreground" Value="#042F2E"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,9"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="BtnBorder" Property="Background" Value="#5EEAD4"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="BtnBorder" Property="Background" Value="#14B8A6"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryBtnStyle" TargetType="Button">
            <Setter Property="Background" Value="#14FFFFFF"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="16,9"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="BtnBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="BtnBorder" Property="Background" Value="#24FFFFFF"/></Trigger>
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

        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="24,14">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="$ToolName" FontSize="17" FontWeight="SemiBold" Foreground="#F8FAFC"/>
                <Border Margin="10,0,0,0" Background="#1A2DD4BF" BorderBrush="#402DD4BF" BorderThickness="1" CornerRadius="4" Padding="8,2" VerticalAlignment="Center">
                    <TextBlock Text="v1.0.0" FontSize="11" FontWeight="SemiBold" Foreground="#2DD4BF"/>
                </Border>
            </StackPanel>
        </Border>

        <Grid Grid.Row="1" Margin="24,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="$Synopsis" Foreground="#CBD5E1" FontSize="13" TextWrapping="Wrap" Margin="0,0,0,14"/>
            <Border Grid.Row="1" Background="#CC050A14" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="12">
                <TextBox x:Name="TxtLogConsole"
                         Background="Transparent"
                         Foreground="#CBD5E1"
                         BorderThickness="0"
                         FontFamily="Cascadia Mono, Consolas"
                         FontSize="12"
                         IsReadOnly="True"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         Text="Ready."/>
            </Border>
        </Grid>

        <Border Grid.Row="2" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" VerticalAlignment="Center" Foreground="#94A3B8" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,0,16,0" Text="Logs: C:\Logs\$ToolName\$ToolName.log"/>
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

    `$txtLogConsole = `$window.FindName('TxtLogConsole')
    `$btnClose = `$window.FindName('BtnClose')
    `$btnRun = `$window.FindName('BtnRun')

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
    Set-Content -Path $scriptPath -Value $psContent -Encoding UTF8
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
    Set-Content -Path $gitignorePath -Value "Output/`r`n*.intunewin`r`n" -Encoding UTF8
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
    Set-Content -Path $readmePath -Value $readmeContent -Encoding UTF8
}

Write-Output "Scaffolded wrapped PowerShell tool at: $targetDir"
