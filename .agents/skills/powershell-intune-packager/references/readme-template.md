# PowerShell Tool `README.md` & STA Reference (`references/readme-template.md`)

## 1. WPF Single-Threaded Apartment (`STA`) Check
```powershell
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
    exit
}
```

## 2. Required `README.md` Section Order
1. `# <Tool Name> Guide` + immediate 1–2 sentence overview paragraph (parsed by `sync_baselines.cjs` into `metadata.json`).
2. `## Overview` + `### Key Features`
3. `## Prerequisites`
4. `## Walkthrough & Usage Guide` (`### 1. Step-by-Step Instructions`, optional `### 2. Intune Portal Upload Parameters` table & `### 3. Step-by-Step Intune Admin Center Setup`, `### Logging & Diagnostics`, `### How Detection & Installation Work`)
5. `## Fast Execute`:
   - **Single-File Remote Script**:
     ````markdown
     > [!TIP]
     > **Short Branded URL (phnx.it):**
     > ```powershell
     > irm https://phnx.it/<CODE> | iex
     > ```
     > **Full Transparent GitHub Command:**
     > ```powershell
     > iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/<Category>/<ToolName>/<ToolName>.ps1")
     > ```
     ````
   - **Multi-File / GUI Package**:
     ````markdown
     > [!WARNING]
     > **Bundled Package Notice:** Multi-file package; clone/download locally rather than `irm | iex`.
     >
     > [!TIP]
     > Double-click `Start-Gui.bat` or run `powershell.exe -ExecutionPolicy Bypass -File .\<ToolName>.ps1`
     ````

## 3. Canonical Intune Win32 Deployment `.txt` Output Template
Whenever an Intune Win32 packager generates a `.intunewin` file and custom detection `.ps1` script, it MUST also generate a deployment instructions `.txt` file matching the exact step-by-step order of the Microsoft Intune "Add Windows app (Win32)" wizard, with the copy-pasteable `Description` box listing `C:\Logs\<PackageFolder>\` and all log files inside that folder:

```text
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : <dd/MM/yy HH:mm:ss>
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
<PackageFileName>.intunewin


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
<App Display Name>

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
<1-2 sentence summary of what the package deploys/configures>
All runtime logs are saved to C:\Logs\<PackageFolder>\

PACKAGE DETAILS:
- <Key Detail 1> : <Value>
- <Key Detail 2> : <Value>
- Author         : JP
- App Version    : <Version>
- Created        : <dd/MM/yy>

WORKSTATION LOG LOCATION:
C:\Logs\<PackageFolder>\

LOG FILES IN THIS FOLDER:
- Install Log   : C:\Logs\<PackageFolder>\<InstallLog>.log
- Detection Log : C:\Logs\<PackageFolder>\<DetectionLog>.log
- Uninstall Log : C:\Logs\<PackageFolder>\<UninstallLog>.log
======================================================================

Publisher:
<PublisherName>

App Version:
<Version>


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\<InstallScript>.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\<UninstallScript>.ps1

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
<DetectionScript>.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No
========================================================================
```

## 4. Canonical Uniform Obsidian Teal + Figtree GUI Templates

### A. Self-Contained WPF Helper (`Initialize-StandardGuiWindow`)
Embed this inside standalone WPF `.ps1` tools (never rely on relative `Import-Module`, so `irm | iex` and `Restricted` execution policies still load `Figtree` and size properly):

```powershell
function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,

        [switch]$PreserveSize
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

    if (-not $PreserveSize) {
        $workArea = [System.Windows.SystemParameters]::WorkArea
        $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
        $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

        $Window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
        $Window.Width     = $targetWidth
        $Window.Height    = $targetHeight
        $Window.MinWidth  = [Math]::Min(860, [int]$targetWidth)
        $Window.MinHeight = [Math]::Min(640, [int]$targetHeight)
        $Window.Left      = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
        $Window.Top       = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
        $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
    }
}
```

### B. Self-Contained WinForms Helper (`Initialize-StandardWinForm`)
Keep `$Script:UiFontCollection` at `$Script:` scope so garbage collection does not release `Figtree` font handles while the form is open:

```powershell
$Script:UiFontCollection = New-Object System.Drawing.Text.PrivateFontCollection
$Script:UiFigtreeFamily = $null

function New-UiFont {
    [CmdletBinding()]
    param(
        [float]$Size = 9.5,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular,
        [switch]$Monospace
    )
    if ($Monospace) {
        foreach ($monoName in @('Cascadia Mono', 'Consolas', 'Courier New')) {
            try { return New-Object System.Drawing.Font($monoName, $Size, $Style) } catch {}
        }
    }
    if ($null -ne $Script:UiFigtreeFamily) {
        try {
            if ($Script:UiFigtreeFamily.IsStyleAvailable($Style)) {
                return New-Object System.Drawing.Font($Script:UiFigtreeFamily, $Size, $Style)
            } elseif ($Script:UiFigtreeFamily.IsStyleAvailable([System.Drawing.FontStyle]::Regular)) {
                return New-Object System.Drawing.Font($Script:UiFigtreeFamily, $Size, [System.Drawing.FontStyle]::Regular)
            }
        } catch {}
    }
    return New-Object System.Drawing.Font('Segoe UI', $Size, $Style)
}

function Initialize-StandardWinForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $resolvedFontDir = $null
    $searchRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { Split-Path -Path $PSCommandPath -Parent } else { $null }
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
            if (-not (Test-Path -Path $cacheDir)) { New-Item -Path $cacheDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
            $allCached = $true
            foreach ($ff in $fontFiles) {
                if (-not (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath $ff) -PathType Leaf)) { $allCached = $false; break }
            }
            if (-not $allCached) {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    foreach ($ff in $fontFiles) {
                        $dest = Join-Path -Path $cacheDir -ChildPath $ff
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/$ff" -OutFile $dest -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
                        }
                    }
                } finally { $ProgressPreference = $prevProgress }
            }
            if (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) { $resolvedFontDir = $cacheDir }
        } catch {}
    }

    if ($resolvedFontDir) {
        foreach ($ff in $fontFiles) {
            $fullFontPath = Join-Path -Path $resolvedFontDir -ChildPath $ff
            if (Test-Path -Path $fullFontPath -PathType Leaf) {
                try { $Script:UiFontCollection.AddFontFile($fullFontPath) } catch {}
            }
        }
        if ($Script:UiFontCollection.Families.Count -gt 0) {
            foreach ($fam in $Script:UiFontCollection.Families) {
                if ($fam.Name -eq 'Figtree') { $Script:UiFigtreeFamily = $fam; break }
            }
            if ($null -eq $Script:UiFigtreeFamily) { $Script:UiFigtreeFamily = $Script:UiFontCollection.Families[0] }
        }
    }

    $workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $targetWidth  = [int][Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [int][Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Form.MaximizeBox     = $true
    $Form.Size            = New-Object System.Drawing.Size($targetWidth, $targetHeight)
    $Form.MinimumSize     = New-Object System.Drawing.Size([Math]::Min(860, $targetWidth), [Math]::Min(640, $targetHeight))
    $Form.Location        = New-Object System.Drawing.Point(
        [int][Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2)),
        [int][Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    )
    $Form.BackColor       = [System.Drawing.Color]::FromArgb(9, 17, 34)
    $Form.ForeColor       = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $Form.Font            = New-UiFont -Size 9.5
}
```

### C. Starter Obsidian Teal WPF Shell (`XAML`)
Do NOT put `FontFamily="Segoe UI*"` on `<Window>` or child styles so `$Window.FontFamily` (`Figtree`) inherits cleanly everywhere (use `FontFamily="Cascadia Mono, Consolas"` strictly for log consoles, paths, and numeric readouts):

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="<Tool Display Name> v1.0.0"
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
        <!-- Band 1: Tool-Only Header -->
        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="24,14">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="<Tool Display Name>" FontSize="17" FontWeight="SemiBold" Foreground="#F8FAFC"/>
                <Border Margin="10,0,0,0" Background="#1A2DD4BF" BorderBrush="#402DD4BF" BorderThickness="1" CornerRadius="4" Padding="8,2" VerticalAlignment="Center">
                    <TextBlock Text="v1.0.0" FontSize="11" FontWeight="SemiBold" Foreground="#2DD4BF"/>
                </Border>
            </StackPanel>
        </Border>
        <!-- Band 2: Flush Workspace -->
        <Grid Grid.Row="1"/>
        <!-- Band 3: Single-Row Footer (Left Hint | Right Horizontal Cluster) -->
        <Border Grid.Row="2" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" VerticalAlignment="Center" Foreground="#94A3B8" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,0,16,0" Text="Logs: C:\Logs\<ToolName>\"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnClose" Style="{StaticResource SecondaryBtnStyle}" Content="Close" Margin="0,0,10,0"/>
                    <Button x:Name="BtnRun" Style="{StaticResource PrimaryCtaStyle}" Content="Run"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
```
