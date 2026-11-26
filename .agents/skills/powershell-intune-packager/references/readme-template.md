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

### A. Self-Contained WPF Helpers (`Win11ChromeHelper`, `Get-SystemIsDarkTheme`, `Set-StandardGuiTheme`, `Initialize-StandardGuiWindow`)
Embed these inside standalone WPF `.ps1` tools (never rely on relative `Import-Module`, so `irm | iex` and `Restricted` execution policies still paint Windows 11 chrome, detect system Light/Dark mode, load `Figtree`, and size properly):

```powershell
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
        $Window.Width      = $targetWidth
        $Window.Height     = $targetHeight
        $Window.MinWidth   = [Math]::Min(860, [int]$targetWidth)
        $Window.MinHeight  = [Math]::Min(640, [int]$targetHeight)
        $Window.Left       = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
        $Window.Top        = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
        $Window.ResizeMode = [System.Windows.ResizeMode]::CanResize
    }
}
```

### B. Self-Contained WinForms Helpers (`Win11ChromeHelper`, `Get-SystemIsDarkTheme`, `Initialize-StandardWinForm`)
Keep `$Script:UiFontCollection` at `$Script:` scope so garbage collection does not release `Figtree` font handles while the form is open, detect Windows system Dark/Light mode at startup, and wire `HandleCreated` to paint the Windows 11 titlebar and border via `Win11ChromeHelper`:

```powershell
# Include Win11ChromeHelper and Get-SystemIsDarkTheme from Section 4A above.

$Script:UiFontCollection = New-Object System.Drawing.Text.PrivateFontCollection
$Script:UiFigtreeFamily = $null
$Script:IsDarkTheme = Get-SystemIsDarkTheme

if ($Script:IsDarkTheme) {
    $clrCanvas    = [System.Drawing.Color]::FromArgb(9, 17, 34)     # #091122
    $clrSurface   = [System.Drawing.Color]::FromArgb(13, 25, 46)    # #0D192E
    $clrInput     = [System.Drawing.Color]::FromArgb(11, 21, 40)    # #0B1528
    $clrLog       = [System.Drawing.Color]::FromArgb(5, 10, 20)     # #050A14
    $clrChrome    = [System.Drawing.Color]::FromArgb(7, 13, 25)     # #070D19
    $clrBorder    = [System.Drawing.Color]::FromArgb(38, 53, 77)    # #26354D
    $clrAccent    = [System.Drawing.Color]::FromArgb(45, 212, 191)  # #2DD4BF
    $clrAccentFg  = [System.Drawing.Color]::FromArgb(4, 47, 46)     # #042F2E
    $clrSecBtn    = [System.Drawing.Color]::FromArgb(19, 34, 58)    # #13223A
    $clrTextPrim  = [System.Drawing.Color]::FromArgb(248, 250, 252) # #F8FAFC
    $clrTextSec   = [System.Drawing.Color]::FromArgb(203, 213, 225) # #CBD5E1
    $clrTextMuted = [System.Drawing.Color]::FromArgb(148, 163, 184) # #94A3B8
    $clrOk        = [System.Drawing.Color]::FromArgb(52, 211, 153)  # #34D399
    $clrErr       = [System.Drawing.Color]::FromArgb(248, 113, 113) # #F87171
} else {
    $clrCanvas    = [System.Drawing.Color]::FromArgb(248, 249, 255) # #F8F9FF
    $clrSurface   = [System.Drawing.Color]::FromArgb(255, 255, 255) # #FFFFFF
    $clrInput     = [System.Drawing.Color]::FromArgb(248, 250, 252) # #F8FAFC
    $clrLog       = [System.Drawing.Color]::FromArgb(241, 245, 249) # #F1F5F9
    $clrChrome    = [System.Drawing.Color]::FromArgb(255, 255, 255) # #FFFFFF
    $clrBorder    = [System.Drawing.Color]::FromArgb(203, 213, 225) # #CBD5E1
    $clrAccent    = [System.Drawing.Color]::FromArgb(37, 99, 235)   # #2563EB
    $clrAccentFg  = [System.Drawing.Color]::FromArgb(255, 255, 255) # #FFFFFF
    $clrSecBtn    = [System.Drawing.Color]::FromArgb(255, 255, 255) # #FFFFFF
    $clrTextPrim  = [System.Drawing.Color]::FromArgb(15, 23, 42)    # #0F172A
    $clrTextSec   = [System.Drawing.Color]::FromArgb(51, 65, 85)    # #334155
    $clrTextMuted = [System.Drawing.Color]::FromArgb(100, 116, 139) # #64748B
    $clrOk        = [System.Drawing.Color]::FromArgb(5, 150, 105)   # #059669
    $clrErr       = [System.Drawing.Color]::FromArgb(220, 38, 38)   # #DC2626
}

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
        [System.Windows.Forms.Form]$Form,

        [bool]$IsDark = $Script:IsDarkTheme
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
    if ($IsDark) {
        $Form.BackColor = [System.Drawing.Color]::FromArgb(9, 17, 34)
        $Form.ForeColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    } else {
        $Form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 255)
        $Form.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    }
    $Form.Font = New-UiFont -Size 9.5

    $chromeDark = $IsDark
    $Form.Add_HandleCreated({
        try {
            if ($Form.Handle -ne [IntPtr]::Zero -and ('Win11ChromeHelper' -as [type])) {
                [Win11ChromeHelper]::ApplyWindowChrome($Form.Handle, $chromeDark)
            }
        } catch {}
    })
}
```

### C. Starter Dynamic Light/Dark WPF Shell (`XAML` + Lifecycle Wiring)
Do NOT put `FontFamily="Segoe UI*"` on `<Window>` or child styles so `$Window.FontFamily` (`Figtree`) inherits cleanly everywhere (use `FontFamily="Cascadia Mono, Consolas"` strictly for log consoles, paths, and numeric readouts). Always use `Background="#091122"` fallback on `<Window>` and assign `$Window.Background = $canvasBrush` in `Set-StandardGuiTheme` (never `{DynamicResource BrushCanvas}` on `<Window>`):

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="<Tool Display Name> v1.0.0"
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
        <!-- Band 1: Tool-Only Header + System Theme Toggle -->
        <Border Grid.Row="0" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,0,0,1" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="<Tool Display Name>" FontSize="17" FontWeight="SemiBold" Foreground="{DynamicResource BrushTextPrimary}"/>
                    <Border Margin="10,0,0,0" Background="{DynamicResource BrushAccentTint}" BorderBrush="{DynamicResource BrushAccentBorder}" BorderThickness="1" CornerRadius="4" Padding="8,2" VerticalAlignment="Center">
                        <TextBlock Text="v1.0.0" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource BrushAccent}"/>
                    </Border>
                </StackPanel>
                <Button x:Name="BtnThemeToggle" Grid.Column="1" Style="{StaticResource SecondaryBtnStyle}" ToolTip="Follows Windows system light/dark theme automatically; click to toggle.">
                    <TextBlock x:Name="TxtThemeLabel" Text="Theme: System (Dark)"/>
                </Button>
            </Grid>
        </Border>
        <!-- Band 2: Flush Workspace -->
        <Grid Grid.Row="1"/>
        <!-- Band 3: Single-Row Footer (Left Hint | Right Horizontal Cluster) -->
        <Border Grid.Row="2" Background="{DynamicResource BrushChromeBar}" BorderBrush="{DynamicResource BrushHairline}" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" VerticalAlignment="Center" Foreground="{DynamicResource BrushTextMuted}" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,0,16,0" Text="Logs: C:\Logs\<ToolName>\"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="BtnClose" Style="{StaticResource SecondaryBtnStyle}" Content="Close" Margin="0,0,10,0"/>
                    <Button x:Name="BtnRun" Style="{StaticResource PrimaryCtaStyle}" Content="Run"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
```

Wire theme detection, painted Win11 chrome, and live `UserPreferenceChanged` tracking before `$window.ShowDialog()`:

```powershell
$txtThemeLabel = $window.FindName('TxtThemeLabel')
$btnThemeToggle = $window.FindName('BtnThemeToggle')

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
```
