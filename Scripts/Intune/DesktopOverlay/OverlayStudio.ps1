<#
.SYNOPSIS
    IT Support Overlay Studio - Interactive Design & Package Builder
.DESCRIPTION
    A visual Windows Forms application allowing IT administrators to design,
    customize, preview, test, and package the IT Support Desktop Overlay into an Intune Win32 (.intunewin) file.
#>

# --- In-Memory Guard for Bundled Packages ---
if ([string]::IsNullOrEmpty($MyInvocation.MyCommand.Path) -or $MyInvocation.MyCommand.Path -match "^iex") {
    throw "In-memory execution (via iex / irm) is not supported for bundled packages. OverlayStudio.ps1 relies on local relative files (src/ payload directory and IntuneWinAppUtil.exe). Please download and extract the repository package locally before executing."
}

# --- Resolve Absolute Script Root Folder ---
$Script:RootPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($Script:RootPath)) {
    $Script:RootPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

# --- Load Assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# High-DPI Awareness
if (-not ([System.Management.Automation.PSTypeName]'StudioDpiHelper').Type) {
    $DpiSignature = @"
using System;
using System.Runtime.InteropServices;
public class StudioDpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
    try { Add-Type -TypeDefinition $DpiSignature -ErrorAction Stop } catch {}
}
try { [StudioDpiHelper]::SetProcessDPIAware() | Out-Null } catch {}

# --- Config State ---
$Config = [ordered]@{
    Title        = "IT SUPPORT HELPDESK"
    Phone        = "0800 123 4567"
    Email        = "support@company.com"
    Hours        = "Mon-Fri: 08:30 - 17:00"
    FieldOrder   = "PhoneFirst"
    Position     = "BottomRight"
    Size         = "Small"
    FontFamily   = "Segoe UI"
    AlwaysOnTop  = $false
    Opacity      = 0.92
    AccentColor  = [System.Drawing.Color]::FromArgb(14, 165, 233)
    BgColor      = [System.Drawing.Color]::FromArgb(26, 32, 44)
    TextColor    = [System.Drawing.Color]::FromArgb(241, 245, 249)
    ShowHost     = $true
    ShowUser     = $true
    ShowIP       = $true
    ShowSerial   = $true
}

# --- Obsidian Dark + Bundled Figtree WinForms Helpers ---
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
            try {
                return New-Object System.Drawing.Font($monoName, $Size, $Style)
            } catch {}
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
        foreach ($ff in $fontFiles) {
            $fullFontPath = Join-Path -Path $resolvedFontDir -ChildPath $ff
            if (Test-Path -Path $fullFontPath -PathType Leaf) {
                try {
                    $Script:UiFontCollection.AddFontFile($fullFontPath)
                } catch {}
            }
        }
        if ($Script:UiFontCollection.Families.Count -gt 0) {
            $Script:UiFigtreeFamily = $Script:UiFontCollection.Families | Where-Object { $_.Name -eq 'Figtree' } | Select-Object -First 1
            if ($null -eq $Script:UiFigtreeFamily) {
                $Script:UiFigtreeFamily = $Script:UiFontCollection.Families[0]
            }
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
    $left = [int][Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $top  = [int][Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Form.Location        = New-Object System.Drawing.Point($left, $top)
    $Form.BackColor       = [System.Drawing.Color]::FromArgb(9, 17, 34)
    $Form.ForeColor       = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $Form.Font            = New-UiFont -Size 9.5
}

$clrCanvas    = [System.Drawing.Color]::FromArgb(9, 17, 34)     # #091122
$clrSurface   = [System.Drawing.Color]::FromArgb(13, 25, 46)    # #0D192E
$clrInput     = [System.Drawing.Color]::FromArgb(11, 21, 40)    # #0B1528
$clrLog       = [System.Drawing.Color]::FromArgb(5, 10, 20)     # #050A14
$clrChrome    = [System.Drawing.Color]::FromArgb(7, 13, 25)     # #070D19
$clrBorder    = [System.Drawing.Color]::FromArgb(38, 53, 77)    # #26354D
$clrTeal      = [System.Drawing.Color]::FromArgb(45, 212, 191)  # #2DD4BF
$clrTealText  = [System.Drawing.Color]::FromArgb(4, 47, 46)     # #042F2E
$clrSecBtn    = [System.Drawing.Color]::FromArgb(19, 34, 58)    # #13223A
$clrTextPrim  = [System.Drawing.Color]::FromArgb(248, 250, 252) # #F8FAFC
$clrTextSec   = [System.Drawing.Color]::FromArgb(203, 213, 225) # #CBD5E1
$clrTextMuted = [System.Drawing.Color]::FromArgb(148, 163, 184) # #94A3B8
$clrOk        = [System.Drawing.Color]::FromArgb(52, 211, 153)  # #34D399
$clrErr       = [System.Drawing.Color]::FromArgb(248, 113, 113) # #F87171
$clrDangerBg  = [System.Drawing.Color]::FromArgb(46, 23, 36)
$clrDangerBr  = [System.Drawing.Color]::FromArgb(153, 27, 27)
$clrDangerFg  = [System.Drawing.Color]::FromArgb(252, 165, 165) # #FCA5A5

function New-BorderedPanel {
    param([int]$PaddingAll = 14)
    $outer = New-Object System.Windows.Forms.Panel
    $outer.BackColor = $clrBorder
    $outer.Padding = New-Object System.Windows.Forms.Padding(1)
    $inner = New-Object System.Windows.Forms.Panel
    $inner.Dock = [System.Windows.Forms.DockStyle]::Fill
    $inner.BackColor = $clrSurface
    $inner.Padding = New-Object System.Windows.Forms.Padding($PaddingAll)
    $outer.Controls.Add($inner)
    return @{ Outer = $outer; Inner = $inner }
}

function Set-InputStyle {
    param($Ctrl)
    $Ctrl.BackColor = $clrInput
    $Ctrl.ForeColor = $clrTextPrim
    if ($Ctrl -is [System.Windows.Forms.TextBox]) {
        $Ctrl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    } elseif ($Ctrl -is [System.Windows.Forms.ComboBox]) {
        $Ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
}

# --- Form Setup ---
$studio = New-Object System.Windows.Forms.Form
$studio.Text = "IT Support Overlay Studio v1.0"
Initialize-StandardWinForm -Form $studio

# --- Header ---
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock      = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height    = 64
$headerPanel.BackColor = $clrChrome
$headerPanel.Padding   = New-Object System.Windows.Forms.Padding(20, 10, 20, 10)
$studio.Controls.Add($headerPanel)

$headerBorder = New-Object System.Windows.Forms.Panel
$headerBorder.Dock = [System.Windows.Forms.DockStyle]::Bottom
$headerBorder.Height = 1
$headerBorder.BackColor = $clrBorder
$headerPanel.Controls.Add($headerBorder)

$headerTitle = New-Object System.Windows.Forms.Label
$headerTitle.Text      = "IT Support Overlay Studio"
$headerTitle.Font      = New-UiFont -Size 14.5 -Style ([System.Drawing.FontStyle]::Bold)
$headerTitle.ForeColor = $clrTextPrim
$headerTitle.Location  = New-Object System.Drawing.Point(20, 10)
$headerTitle.AutoSize  = $true
$headerPanel.Controls.Add($headerTitle)

$headerSub = New-Object System.Windows.Forms.Label
$headerSub.Text      = "Customize design, preview live, and export Intune (.intunewin) package"
$headerSub.Font      = New-UiFont -Size 9
$headerSub.ForeColor = $clrTextMuted
$headerSub.Location  = New-Object System.Drawing.Point(20, 35)
$headerSub.AutoSize  = $true
$headerPanel.Controls.Add($headerSub)

# --- Single-Row Footer ---
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$footerPanel.Height = 58
$footerPanel.BackColor = $clrChrome
$studio.Controls.Add($footerPanel)

$footerBorder = New-Object System.Windows.Forms.Panel
$footerBorder.Dock = [System.Windows.Forms.DockStyle]::Top
$footerBorder.Height = 1
$footerBorder.BackColor = $clrBorder
$footerPanel.Controls.Add($footerBorder)

$footerTable = New-Object System.Windows.Forms.TableLayoutPanel
$footerTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$footerTable.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 10)
$footerTable.ColumnCount = 2
$footerTable.RowCount = 1
[void]$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$footerPanel.Controls.Add($footerTable)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblStatus.AutoEllipsis = $true
$lblStatus.ForeColor = $clrTextMuted
$lblStatus.Font      = New-UiFont -Size 9
$lblStatus.Text      = "Ready. Adjust overlay properties and export to Intune."
$footerTable.Controls.Add($lblStatus, 0, 0)

$footerActions = New-Object System.Windows.Forms.FlowLayoutPanel
$footerActions.AutoSize = $true
$footerActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$footerActions.WrapContents = $false
$footerActions.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$footerTable.Controls.Add($footerActions, 1, 0)

$btnOpenOutput = New-Object System.Windows.Forms.Button
$btnOpenOutput.Text      = "Open Output Folder"
$btnOpenOutput.Size      = New-Object System.Drawing.Size(150, 34)
$btnOpenOutput.Margin    = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$btnOpenOutput.BackColor = $clrSecBtn
$btnOpenOutput.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$btnOpenOutput.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenOutput.FlatAppearance.BorderSize = 1
$btnOpenOutput.FlatAppearance.BorderColor = $clrBorder
$btnOpenOutput.Font      = New-UiFont -Size 9 -Style ([System.Drawing.FontStyle]::Bold)
$btnOpenOutput.Cursor    = [System.Windows.Forms.Cursors]::Hand
$footerActions.Controls.Add($btnOpenOutput)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text      = "Export and Build Intune Package"
$btnExport.Size      = New-Object System.Drawing.Size(245, 34)
$btnExport.Margin    = New-Object System.Windows.Forms.Padding(0)
$btnExport.BackColor = $clrTeal
$btnExport.ForeColor = $clrTealText
$btnExport.Font      = New-UiFont -Size 9.5 -Style ([System.Drawing.FontStyle]::Bold)
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnExport.FlatAppearance.BorderSize = 0
$btnExport.Cursor    = [System.Windows.Forms.Cursors]::Hand
$footerActions.Controls.Add($btnExport)

# --- Main 2-Column Split ---
$mainSplit = New-Object System.Windows.Forms.TableLayoutPanel
$mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainSplit.BackColor = $clrCanvas
$mainSplit.Padding = New-Object System.Windows.Forms.Padding(18, 14, 18, 14)
$mainSplit.ColumnCount = 2
$mainSplit.RowCount = 1
[void]$mainSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 52)))
[void]$mainSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 48)))
[void]$mainSplit.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$studio.Controls.Add($mainSplit)
$mainSplit.BringToFront()

$leftWrap = New-BorderedPanel -PaddingAll 14
$leftWrap.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$leftWrap.Outer.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
$mainSplit.Controls.Add($leftWrap.Outer, 0, 0)
$leftPanel = $leftWrap.Inner
$leftPanel.AutoScroll = $true

$rightWrap = New-BorderedPanel -PaddingAll 14
$rightWrap.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$rightWrap.Outer.Margin = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$mainSplit.Controls.Add($rightWrap.Outer, 1, 0)
$rightPanel = $rightWrap.Inner

function Add-SectionLabel ($parent, $text, $yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $text
    $lbl.Font      = New-UiFont -Size 9.5 -Style ([System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = $clrTeal
    $lbl.Location  = New-Object System.Drawing.Point(12, $yPos)
    $lbl.AutoSize  = $true
    $parent.Controls.Add($lbl)
}

# --- LEFT PANEL: CONTROLS ---

# Section 1: Content and Typography
Add-SectionLabel $leftPanel "1. CONTENT & TYPOGRAPHY" 10

# Title Text
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Header Title:"; $lblTitle.ForeColor = $clrTextSec; $lblTitle.Location = New-Object System.Drawing.Point(12, 38); $lblTitle.AutoSize = $true; $leftPanel.Controls.Add($lblTitle)
$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Text = $Config.Title; $txtTitle.Location = New-Object System.Drawing.Point(130, 35); $txtTitle.Size = New-Object System.Drawing.Size(295, 24)
$txtTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $txtTitle; $leftPanel.Controls.Add($txtTitle)

# Phone
$lblPhone = New-Object System.Windows.Forms.Label
$lblPhone.Text = "Support Phone:"; $lblPhone.ForeColor = $clrTextSec; $lblPhone.Location = New-Object System.Drawing.Point(12, 68); $lblPhone.AutoSize = $true; $leftPanel.Controls.Add($lblPhone)
$txtPhone = New-Object System.Windows.Forms.TextBox
$txtPhone.Text = $Config.Phone; $txtPhone.Location = New-Object System.Drawing.Point(130, 65); $txtPhone.Size = New-Object System.Drawing.Size(295, 24)
$txtPhone.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $txtPhone; $leftPanel.Controls.Add($txtPhone)

# Email
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "Support Email:"; $lblEmail.ForeColor = $clrTextSec; $lblEmail.Location = New-Object System.Drawing.Point(12, 98); $lblEmail.AutoSize = $true; $leftPanel.Controls.Add($lblEmail)
$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Text = $Config.Email; $txtEmail.Location = New-Object System.Drawing.Point(130, 95); $txtEmail.Size = New-Object System.Drawing.Size(295, 24)
$txtEmail.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $txtEmail; $leftPanel.Controls.Add($txtEmail)

# Hours (Optional)
$lblHours = New-Object System.Windows.Forms.Label
$lblHours.Text = "Support Hours:"; $lblHours.ForeColor = $clrTextSec; $lblHours.Location = New-Object System.Drawing.Point(12, 128); $lblHours.AutoSize = $true; $leftPanel.Controls.Add($lblHours)
$txtHours = New-Object System.Windows.Forms.TextBox
$txtHours.Text = $Config.Hours; $txtHours.Location = New-Object System.Drawing.Point(130, 125); $txtHours.Size = New-Object System.Drawing.Size(295, 24)
$txtHours.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $txtHours; $leftPanel.Controls.Add($txtHours)

# Font Family Combo
$lblFont = New-Object System.Windows.Forms.Label
$lblFont.Text = "Font Family:"; $lblFont.ForeColor = $clrTextSec; $lblFont.Location = New-Object System.Drawing.Point(12, 158); $lblFont.AutoSize = $true; $leftPanel.Controls.Add($lblFont)
$comboFont = New-Object System.Windows.Forms.ComboBox
$comboFont.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboFont.Items.AddRange(@("Segoe UI", "Consolas", "Tahoma"))
$comboFont.SelectedItem = "Segoe UI"
$comboFont.Location = New-Object System.Drawing.Point(130, 155); $comboFont.Size = New-Object System.Drawing.Size(295, 24)
$comboFont.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $comboFont; $leftPanel.Controls.Add($comboFont)

# Field Order Combo
$lblOrder = New-Object System.Windows.Forms.Label
$lblOrder.Text = "Field Order:"; $lblOrder.ForeColor = $clrTextSec; $lblOrder.Location = New-Object System.Drawing.Point(12, 188); $lblOrder.AutoSize = $true; $leftPanel.Controls.Add($lblOrder)
$comboOrder = New-Object System.Windows.Forms.ComboBox
$comboOrder.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboOrder.Items.AddRange(@("Phone First (Phone then Email)", "Email First (Email then Phone)"))
$comboOrder.SelectedIndex = 0
$comboOrder.Location = New-Object System.Drawing.Point(130, 185); $comboOrder.Size = New-Object System.Drawing.Size(295, 24)
$comboOrder.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $comboOrder; $leftPanel.Controls.Add($comboOrder)

# Section 2: Colors & Theme
Add-SectionLabel $leftPanel "2. THEME COLORS & OPACITY" 222

# Accent Color Picker Button
$btnAccent = New-Object System.Windows.Forms.Button
$btnAccent.Text      = "Accent Color"
$btnAccent.Location  = New-Object System.Drawing.Point(12, 248)
$btnAccent.Size      = New-Object System.Drawing.Size(125, 28)
$btnAccent.BackColor = $Config.AccentColor
$btnAccent.ForeColor = [System.Drawing.Color]::White
$btnAccent.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnAccent.FlatAppearance.BorderSize = 1
$btnAccent.FlatAppearance.BorderColor = $clrBorder
$leftPanel.Controls.Add($btnAccent)

# Background Color Picker Button
$btnBg = New-Object System.Windows.Forms.Button
$btnBg.Text      = "Background"
$btnBg.Location  = New-Object System.Drawing.Point(147, 248)
$btnBg.Size      = New-Object System.Drawing.Size(125, 28)
$btnBg.BackColor = $Config.BgColor
$btnBg.ForeColor = [System.Drawing.Color]::White
$btnBg.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBg.FlatAppearance.BorderSize = 1
$btnBg.FlatAppearance.BorderColor = $clrBorder
$leftPanel.Controls.Add($btnBg)

# Text Color Picker Button
$btnText = New-Object System.Windows.Forms.Button
$btnText.Text      = "Text Color"
$btnText.Location  = New-Object System.Drawing.Point(282, 248)
$btnText.Size      = New-Object System.Drawing.Size(125, 28)
$btnText.BackColor = $Config.TextColor
$btnText.ForeColor = [System.Drawing.Color]::Black
$btnText.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnText.FlatAppearance.BorderSize = 1
$btnText.FlatAppearance.BorderColor = $clrBorder
$leftPanel.Controls.Add($btnText)

# Opacity Slider
$lblOpacity = New-Object System.Windows.Forms.Label
$lblOpacity.Text = "Card Opacity (92%):"; $lblOpacity.ForeColor = $clrTextSec; $lblOpacity.Location = New-Object System.Drawing.Point(12, 290); $lblOpacity.AutoSize = $true; $leftPanel.Controls.Add($lblOpacity)
$trackOpacity = New-Object System.Windows.Forms.TrackBar
$trackOpacity.Minimum = 60; $trackOpacity.Maximum = 100; $trackOpacity.Value = 92
$trackOpacity.BackColor = $clrSurface
$trackOpacity.Location = New-Object System.Drawing.Point(150, 284); $trackOpacity.Size = New-Object System.Drawing.Size(275, 40)
$trackOpacity.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$leftPanel.Controls.Add($trackOpacity)

# Section 3: Placement, Sizing & Z-Order Layer
Add-SectionLabel $leftPanel "3. PLACEMENT, SIZING & Z-ORDER" 334

# Position Combo
$lblPos = New-Object System.Windows.Forms.Label
$lblPos.Text = "Screen Corner:"; $lblPos.ForeColor = $clrTextSec; $lblPos.Location = New-Object System.Drawing.Point(12, 362); $lblPos.AutoSize = $true; $leftPanel.Controls.Add($lblPos)
$comboPos = New-Object System.Windows.Forms.ComboBox
$comboPos.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboPos.Items.AddRange(@("BottomRight", "TopRight", "BottomLeft", "TopLeft"))
$comboPos.SelectedItem = "BottomRight"
$comboPos.Location = New-Object System.Drawing.Point(130, 359); $comboPos.Size = New-Object System.Drawing.Size(295, 24)
$comboPos.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $comboPos; $leftPanel.Controls.Add($comboPos)

# Widget Size Combo
$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Text = "Widget Size:"; $lblSize.ForeColor = $clrTextSec; $lblSize.Location = New-Object System.Drawing.Point(12, 394); $lblSize.AutoSize = $true; $leftPanel.Controls.Add($lblSize)
$comboSize = New-Object System.Windows.Forms.ComboBox
$comboSize.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboSize.Items.AddRange(@("Small (310 x 160 px)", "Medium (380 x 200 px)", "Large (460 x 240 px)"))
$comboSize.SelectedIndex = 0
$comboSize.Location = New-Object System.Drawing.Point(130, 391); $comboSize.Size = New-Object System.Drawing.Size(295, 24)
$comboSize.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $comboSize; $leftPanel.Controls.Add($comboSize)

# Z-Order Layer Combo
$lblLayer = New-Object System.Windows.Forms.Label
$lblLayer.Text = "Z-Order Layer:"; $lblLayer.ForeColor = $clrTextSec; $lblLayer.Location = New-Object System.Drawing.Point(12, 426); $lblLayer.AutoSize = $true; $leftPanel.Controls.Add($lblLayer)
$comboLayer = New-Object System.Windows.Forms.ComboBox
$comboLayer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboLayer.Items.AddRange(@("Desktop Background (Behind Apps)", "Always On Top"))
$comboLayer.SelectedIndex = 0
$comboLayer.Location = New-Object System.Drawing.Point(130, 423); $comboLayer.Size = New-Object System.Drawing.Size(295, 24)
$comboLayer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
Set-InputStyle $comboLayer; $leftPanel.Controls.Add($comboLayer)

# Section 4: System Fields Toggle
Add-SectionLabel $leftPanel "4. VISIBLE DIAGNOSTIC FIELDS" 464

$chkHost = New-Object System.Windows.Forms.CheckBox
$chkHost.Text = "Host Name"; $chkHost.Checked = $true; $chkHost.ForeColor = $clrTextSec; $chkHost.Location = New-Object System.Drawing.Point(14, 492); $chkHost.AutoSize = $true; $leftPanel.Controls.Add($chkHost)

$chkUser = New-Object System.Windows.Forms.CheckBox
$chkUser.Text = "User Name"; $chkUser.Checked = $true; $chkUser.ForeColor = $clrTextSec; $chkUser.Location = New-Object System.Drawing.Point(116, 492); $chkUser.AutoSize = $true; $leftPanel.Controls.Add($chkUser)

$chkIP = New-Object System.Windows.Forms.CheckBox
$chkIP.Text = "IPv4 Address"; $chkIP.Checked = $true; $chkIP.ForeColor = $clrTextSec; $chkIP.Location = New-Object System.Drawing.Point(218, 492); $chkIP.AutoSize = $true; $leftPanel.Controls.Add($chkIP)

$chkSerial = New-Object System.Windows.Forms.CheckBox
$chkSerial.Text = "Serial Number"; $chkSerial.Checked = $true; $chkSerial.ForeColor = $clrTextSec; $chkSerial.Location = New-Object System.Drawing.Point(324, 492); $chkSerial.AutoSize = $true; $leftPanel.Controls.Add($chkSerial)

# --- RIGHT PANEL: LIVE PREVIEW & TEST CONTROLS ---
$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 4
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 64)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 36)))
$rightPanel.Controls.Add($rightLayout)

$lblPreviewHead = New-Object System.Windows.Forms.Label
$lblPreviewHead.Text = "LIVE INTERACTIVE PREVIEW"
$lblPreviewHead.Font = New-UiFont -Size 9.5 -Style ([System.Drawing.FontStyle]::Bold)
$lblPreviewHead.ForeColor = $clrTeal
$lblPreviewHead.AutoSize = $true
$lblPreviewHead.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
$rightLayout.Controls.Add($lblPreviewHead, 0, 0)

# Preview Box Container
$previewOuter = New-Object System.Windows.Forms.Panel
$previewOuter.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewOuter.BackColor = $clrBorder
$previewOuter.Padding = New-Object System.Windows.Forms.Padding(1)
$previewOuter.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$rightLayout.Controls.Add($previewOuter, 0, 1)

$previewBox = New-Object System.Windows.Forms.Panel
$previewBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewBox.BackColor = $clrLog
$previewOuter.Controls.Add($previewBox)

# Embedded Preview Card Form
$previewCard = New-Object System.Windows.Forms.Panel
$previewCard.Size      = New-Object System.Drawing.Size(320, 160)
$previewCard.Location  = New-Object System.Drawing.Point(29, 35)
$previewCard.BackColor = $Config.BgColor
$previewBox.Controls.Add($previewCard)

$previewAccent = New-Object System.Windows.Forms.Panel
$previewAccent.Height    = 4
$previewAccent.Dock      = [System.Windows.Forms.DockStyle]::Top
$previewAccent.BackColor = $Config.AccentColor
$previewCard.Controls.Add($previewAccent)

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Dock      = [System.Windows.Forms.DockStyle]::Fill
$previewLabel.ForeColor = $Config.TextColor
$previewLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$previewLabel.Padding   = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$previewCard.Controls.Add($previewLabel)

# Test Buttons Row
$testRow = New-Object System.Windows.Forms.TableLayoutPanel
$testRow.Dock = [System.Windows.Forms.DockStyle]::Fill
$testRow.AutoSize = $true
$testRow.ColumnCount = 2
$testRow.RowCount = 1
$testRow.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
[void]$testRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 65)))
[void]$testRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 35)))
$rightLayout.Controls.Add($testRow, 0, 2)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text      = "Test Live Overlay on Desktop"
$btnTest.Dock      = [System.Windows.Forms.DockStyle]::Fill
$btnTest.Height    = 36
$btnTest.Margin    = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btnTest.BackColor = $clrSecBtn
$btnTest.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
$btnTest.Font      = New-UiFont -Size 9 -Style ([System.Drawing.FontStyle]::Bold)
$btnTest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTest.FlatAppearance.BorderSize = 1
$btnTest.FlatAppearance.BorderColor = $clrBorder
$btnTest.Cursor    = [System.Windows.Forms.Cursors]::Hand
$testRow.Controls.Add($btnTest, 0, 0)

$btnCloseTest = New-Object System.Windows.Forms.Button
$btnCloseTest.Text      = "Stop Overlay"
$btnCloseTest.Dock      = [System.Windows.Forms.DockStyle]::Fill
$btnCloseTest.Height    = 36
$btnCloseTest.Margin    = New-Object System.Windows.Forms.Padding(0)
$btnCloseTest.BackColor = $clrDangerBg
$btnCloseTest.ForeColor = $clrDangerFg
$btnCloseTest.Font      = New-UiFont -Size 9 -Style ([System.Drawing.FontStyle]::Bold)
$btnCloseTest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCloseTest.FlatAppearance.BorderSize = 1
$btnCloseTest.FlatAppearance.BorderColor = $clrDangerBr
$btnCloseTest.Cursor    = [System.Windows.Forms.Cursors]::Hand
$testRow.Controls.Add($btnCloseTest, 1, 0)

# Bundle Artifacts Console
$txtOutputInfo = New-Object System.Windows.Forms.TextBox
$txtOutputInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtOutputInfo.Multiline = $true
$txtOutputInfo.ReadOnly = $true
$txtOutputInfo.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtOutputInfo.BackColor = $clrLog
$txtOutputInfo.ForeColor = $clrTextSec
$txtOutputInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtOutputInfo.Font = New-UiFont -Size 8.5 -Monospace
$txtOutputInfo.Text = @(
    "EXPORTED BUNDLE ARTIFACTS (output\):",
    " - Install-ITOverlay.intunewin",
    " - Detect-ITOverlay.ps1",
    " - Intune-Deployment-Instructions.txt",
    "",
    "Saves defaults to src\Start-ITOverlay.ps1 and syncs version in src\Detect-ITOverlay.ps1."
) -join "`r`n"
$rightLayout.Controls.Add($txtOutputInfo, 0, 3)

# Function to Update Preview Card Text & Colors
function Update-Preview {
    $Config.Title       = $txtTitle.Text
    $Config.Phone       = $txtPhone.Text
    $Config.Email       = $txtEmail.Text
    $Config.Hours       = $txtHours.Text
    $Config.FieldOrder  = if ($comboOrder.SelectedIndex -eq 1) { "EmailFirst" } else { "PhoneFirst" }
    $Config.FontFamily  = if ($comboFont.SelectedItem) { $comboFont.SelectedItem } else { "Segoe UI" }
    $Config.Size        = switch ($comboSize.SelectedIndex) {
        1 { "Medium" }
        2 { "Large" }
        default { "Small" }
    }
    $Config.Opacity     = ($trackOpacity.Value / 100.0)
    $Config.Position    = $comboPos.SelectedItem
    $Config.AlwaysOnTop = ($comboLayer.SelectedIndex -eq 1)
    $Config.ShowHost    = $chkHost.Checked
    $Config.ShowUser    = $chkUser.Checked
    $Config.ShowIP      = $chkIP.Checked
    $Config.ShowSerial  = $chkSerial.Checked

    $lblOpacity.Text = "Card Opacity ($($trackOpacity.Value)%):"

    # Color & Card Updates
    $previewCard.BackColor   = $Config.BgColor
    $previewAccent.BackColor = $Config.AccentColor
    $previewLabel.ForeColor  = $Config.TextColor

    switch ($Config.Size) {
        "Medium" {
            $cardW = 350; $cardH = 180
            $pFontSize = 10.0
            $dashSep = "----------------------------------------------------"
        }
        "Large" {
            $cardW = 366; $cardH = 200
            $pFontSize = 11.0
            $dashSep = "------------------------------------------------------------------"
        }
        default { # Small
            $cardW = 320; $cardH = 160
            $pFontSize = 9.0
            $dashSep = "-----------------------------------------------"
        }
    }

    $previewCard.Size = New-Object System.Drawing.Size($cardW, $cardH)
    $px = [Math]::Max(8, [int](($previewBox.ClientSize.Width - $cardW) / 2))
    $py = [Math]::Max(8, [int](($previewBox.ClientSize.Height - $cardH) / 2))
    $previewCard.Location = New-Object System.Drawing.Point($px, $py)

    $previewLabel.Font = try { New-Object System.Drawing.Font($Config.FontFamily, $pFontSize) } catch { New-Object System.Drawing.Font("Segoe UI", $pFontSize) }

    # Build Preview String
    $lines = @()
    $lines += $Config.Title.ToUpper()
    $lines += $dashSep

    if ($Config.FieldOrder -eq "EmailFirst") {
        if (-not [string]::IsNullOrWhiteSpace($Config.Email)) { $lines += "Email:  $($Config.Email)" }
        if (-not [string]::IsNullOrWhiteSpace($Config.Phone)) { $lines += "Phone:  $($Config.Phone)" }
    } else {
        if (-not [string]::IsNullOrWhiteSpace($Config.Phone)) { $lines += "Phone:  $($Config.Phone)" }
        if (-not [string]::IsNullOrWhiteSpace($Config.Email)) { $lines += "Email:  $($Config.Email)" }
    }
    if (-not [string]::IsNullOrWhiteSpace($Config.Hours)) {
        $lines += "Hours:  $($Config.Hours)"
    }
    $lines += ""

    $sysFields = @()
    if ($Config.ShowHost)   { $sysFields += "Host: $env:COMPUTERNAME" }
    if ($Config.ShowUser)   { $sysFields += "User: $env:USERNAME" }
    if ($sysFields.Count -gt 0) { $lines += ($sysFields -join "  |  ") }

    $detailFields = @()
    if ($Config.ShowIP)     { $detailFields += "IP: 192.168.1.50" }
    if ($Config.ShowSerial) { $detailFields += "Serial: ABC12345" }
    if ($detailFields.Count -gt 0) { $lines += ($detailFields -join "  |  ") }

    $previewLabel.Text = ($lines -join "`n")
}

$previewBox.Add_Resize({ Update-Preview })

# Attach Live Update Events
$txtTitle.Add_TextChanged({ Update-Preview })
$txtPhone.Add_TextChanged({ Update-Preview })
$txtEmail.Add_TextChanged({ Update-Preview })
$txtHours.Add_TextChanged({ Update-Preview })
$comboFont.Add_SelectedIndexChanged({ Update-Preview })
$comboOrder.Add_SelectedIndexChanged({ Update-Preview })
$comboSize.Add_SelectedIndexChanged({ Update-Preview })
$trackOpacity.Add_ValueChanged({ Update-Preview })
$comboPos.Add_SelectedIndexChanged({ Update-Preview })
$comboLayer.Add_SelectedIndexChanged({ Update-Preview })
$chkHost.Add_CheckedChanged({ Update-Preview })
$chkUser.Add_CheckedChanged({ Update-Preview })
$chkIP.Add_CheckedChanged({ Update-Preview })
$chkSerial.Add_CheckedChanged({ Update-Preview })

# Color Pickers
$btnAccent.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = $Config.AccentColor
    if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Config.AccentColor = $cd.Color
        $btnAccent.BackColor = $cd.Color
        Update-Preview
    }
})

$btnBg.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = $Config.BgColor
    if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Config.BgColor = $cd.Color
        $btnBg.BackColor = $cd.Color
        Update-Preview
    }
})

$btnText.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = $Config.TextColor
    if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Config.TextColor = $cd.Color
        $btnText.BackColor = $cd.Color
        Update-Preview
    }
})

# Helper function to save defaults directly into src/Start-ITOverlay.ps1
function Save-OverlayDefault {
    $srcFile = Join-Path -Path $Script:RootPath -ChildPath "src\Start-ITOverlay.ps1"
    if (Test-Path -Path $srcFile) {
        $content = Get-Content -Path $srcFile -Raw
        
        $accentHex = "#{0:X2}{1:X2}{2:X2}" -f $Config.AccentColor.R, $Config.AccentColor.G, $Config.AccentColor.B
        $bgHex     = "#{0:X2}{1:X2}{2:X2}" -f $Config.BgColor.R, $Config.BgColor.G, $Config.BgColor.B
        $textHex   = "#{0:X2}{1:X2}{2:X2}" -f $Config.TextColor.R, $Config.TextColor.G, $Config.TextColor.B

        $content = $content -replace '\[string\]\$Title\s*=\s*"[^"]*"', ('[string]$Title = "{0}"' -f $Config.Title)
        $content = $content -replace '\[string\]\$SupportPhone\s*=\s*"[^"]*"', ('[string]$SupportPhone = "{0}"' -f $Config.Phone)
        $content = $content -replace '\[string\]\$SupportEmail\s*=\s*"[^"]*"', ('[string]$SupportEmail = "{0}"' -f $Config.Email)
        $content = $content -replace '\[string\]\$SupportHours\s*=\s*"[^"]*"', ('[string]$SupportHours = "{0}"' -f $Config.Hours)
        $content = $content -replace '\[string\]\$FieldOrder\s*=\s*"[^"]*"', ('[string]$FieldOrder = "{0}"' -f $Config.FieldOrder)
        $content = $content -replace '\[string\]\$Position\s*=\s*"[^"]*"', ('[string]$Position = "{0}"' -f $Config.Position)
        $content = $content -replace '\[string\]\$Size\s*=\s*"[^"]*"', ('[string]$Size = "{0}"' -f $Config.Size)
        $content = $content -replace '\[string\]\$FontFamily\s*=\s*"[^"]*"', ('[string]$FontFamily = "{0}"' -f $Config.FontFamily)
        $content = $content -replace '\[string\]\$AccentColorHex\s*=\s*"[^"]*"', ('[string]$AccentColorHex = "{0}"' -f $accentHex)
        $content = $content -replace '\[string\]\$BgColorHex\s*=\s*"[^"]*"', ('[string]$BgColorHex = "{0}"' -f $bgHex)
        $content = $content -replace '\[string\]\$TextColorHex\s*=\s*"[^"]*"', ('[string]$TextColorHex = "{0}"' -f $textHex)
        
        $boolHostStr   = if ($Config.ShowHost) { '$true' } else { '$false' }
        $boolUserStr   = if ($Config.ShowUser) { '$true' } else { '$false' }
        $boolIPStr     = if ($Config.ShowIP) { '$true' } else { '$false' }
        $boolSerialStr = if ($Config.ShowSerial) { '$true' } else { '$false' }

        $content = $content -replace '\[bool\]\$ShowHost\s*=\s*\$(true|false)', ('[bool]$ShowHost = {0}' -f $boolHostStr)
        $content = $content -replace '\[bool\]\$ShowUser\s*=\s*\$(true|false)', ('[bool]$ShowUser = {0}' -f $boolUserStr)
        $content = $content -replace '\[bool\]\$ShowIP\s*=\s*\$(true|false)', ('[bool]$ShowIP = {0}' -f $boolIPStr)
        $content = $content -replace '\[bool\]\$ShowSerial\s*=\s*\$(true|false)', ('[bool]$ShowSerial = {0}' -f $boolSerialStr)

        $newBuildVersion = "1.0." + (Get-Date -Format "ddMMyy.HHmm")
        $content = $content -replace '\[string\]\$BuildVersion\s*=\s*"[^"]*"', ('[string]$BuildVersion = "{0}"' -f $newBuildVersion)

        Set-Content -Path $srcFile -Value $content -Encoding UTF8

        # Synchronize TargetVersion in src/Detect-ITOverlay.ps1
        $detectFile = Join-Path -Path $Script:RootPath -ChildPath "src\Detect-ITOverlay.ps1"
        if (Test-Path -Path $detectFile) {
            $detectContent = Get-Content -Path $detectFile -Raw
            $detectContent = $detectContent -replace '\[string\]\$TargetVersion\s*=\s*"[^"]*"', ('[string]$TargetVersion = "{0}"' -f $newBuildVersion)
            Set-Content -Path $detectFile -Value $detectContent -Encoding UTF8
        }
    }
}

# --- ACTION HANDLERS ---

$btnOpenOutput.Add_Click({
    $outputFolder = Join-Path -Path $Script:RootPath -ChildPath "output"
    if (-not (Test-Path -Path $outputFolder)) {
        try { New-Item -ItemType Directory -Path $outputFolder -Force -ErrorAction Stop | Out-Null } catch {}
    }
    if (Test-Path -Path $outputFolder) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$outputFolder`""
    }
})

$btnCloseTest.Add_Click({
    try {
        $procs = try { Get-CimInstance -ClassName Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop } catch { @() }
        $procs | 
            Where-Object { $_.CommandLine -like "*Start-ITOverlay.ps1*" } | 
            ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {} }
        
        $lblStatus.Text = "Live test overlay stopped."
        $lblStatus.ForeColor = $clrErr
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error stopping test overlay:`n`n$_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnTest.Add_Click({
    try {
        # Close any previous live test overlay processes first
        try {
            $procs = try { Get-CimInstance -ClassName Win32_Process -Filter "Name='powershell.exe'" -ErrorAction Stop } catch { @() }
            $procs | 
                Where-Object { $_.CommandLine -like "*Start-ITOverlay.ps1*" } | 
                ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {} }
        } catch {}

        Update-Preview
        $ScriptPath = Join-Path -Path $Script:RootPath -ChildPath "src\Start-ITOverlay.ps1"
        $accentHex = "#{0:X2}{1:X2}{2:X2}" -f $Config.AccentColor.R, $Config.AccentColor.G, $Config.AccentColor.B
        $bgHex     = "#{0:X2}{1:X2}{2:X2}" -f $Config.BgColor.R, $Config.BgColor.G, $Config.BgColor.B
        $textHex   = "#{0:X2}{1:X2}{2:X2}" -f $Config.TextColor.R, $Config.TextColor.G, $Config.TextColor.B

        $hostVal   = if ($Config.ShowHost) { 1 } else { 0 }
        $userVal   = if ($Config.ShowUser) { 1 } else { 0 }
        $ipVal     = if ($Config.ShowIP) { 1 } else { 0 }
        $serialVal = if ($Config.ShowSerial) { 1 } else { 0 }

        $cleanTitle = $Config.Title.Replace("'", "''")
        $cleanPhone = $Config.Phone.Replace("'", "''")
        $cleanEmail = $Config.Email.Replace("'", "''")
        $cleanHours = $Config.Hours.Replace("'", "''")

        $cmdStr = "& '$ScriptPath' -Title '$cleanTitle' -SupportPhone '$cleanPhone' -SupportEmail '$cleanEmail' -SupportHours '$cleanHours' -FieldOrder '$($Config.FieldOrder)' -Position '$($Config.Position)' -Size '$($Config.Size)' -FontFamily '$($Config.FontFamily)' -AccentColorHex '$accentHex' -BgColorHex '$bgHex' -TextColorHex '$textHex' -ShowHost $hostVal -ShowUser $userVal -ShowIP $ipVal -ShowSerial $serialVal"
        if ($Config.AlwaysOnTop) { $cmdStr += " -AlwaysOnTop" }

        Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$cmdStr`"" -WindowStyle Hidden
        $lblStatus.Text = "Live test overlay active on desktop."
        $lblStatus.ForeColor = $clrTeal
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error testing live overlay:`n`n$_", "Test Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnExport.Add_Click({
    try {
        Update-Preview
        $lblStatus.Text = "Saving custom defaults and building Intune package..."
        $lblStatus.ForeColor = $clrTeal
        $studio.Refresh()

        # 1. Update src/Start-ITOverlay.ps1 defaults
        Save-OverlayDefault

        # 2. Run Build-IntunePackage.ps1
        $builderScript = Join-Path -Path $Script:RootPath -ChildPath "Build-IntunePackage.ps1"
        $p = Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$builderScript`"" -PassThru -NoNewWindow -Wait

        if ($p.ExitCode -eq 0) {
            $outputFolder = Join-Path -Path $Script:RootPath -ChildPath "output"
            $outFile = Join-Path -Path $outputFolder -ChildPath "Install-ITOverlay.intunewin"
            if (Test-Path -Path $outFile) {
                $lblStatus.Text = "SUCCESS! Package, Detection Script & Instructions exported to output/ folder."
                $lblStatus.ForeColor = $clrOk
                [System.Windows.Forms.MessageBox]::Show("Custom design saved & Intune Package compiled successfully!`n`nOutput Folder Contents:`n  • Install-ITOverlay.intunewim (App Package)`n  • Detect-ITOverlay.ps1 (Intune Custom Detection Script)`n  • Intune-Deployment-Instructions.txt (Install / Uninstall Commands)`n`nOutput Folder Location:`n" + $outputFolder, "Build Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            $lblStatus.Text = "Error compiling .intunewim package. Exit code: " + $p.ExitCode
            $lblStatus.ForeColor = $clrErr
        }
    } catch {
        $lblStatus.Text = "Error during export: $_"
        $lblStatus.ForeColor = $clrErr
        [System.Windows.Forms.MessageBox]::Show("Error during export:`n`n$_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Initial Render
Update-Preview

# Show Application
[System.Windows.Forms.Application]::Run($studio)
