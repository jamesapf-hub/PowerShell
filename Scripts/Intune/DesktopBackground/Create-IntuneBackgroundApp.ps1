<#
.SYNOPSIS
    GUI utility for building, staging, and packaging corporate Desktop Background and Lock Screen deployments for Intune.
.DESCRIPTION
    This script launches a Windows Forms user interface to configure corporate branding assets. It supports local image files or remote URLs, and generates staging folders, installation scripts (Install.ps1, ApplyBackground.ps1, DownloadBackgrounds.ps1), custom detection scripts, and packs everything into a ready-to-upload .intunewin bundle.
.PARAMETER DebugMode
    Launches the packaging process with verbose execution logs.
.NOTES
    Script Name   : Create-IntuneBackgroundApp.ps1
    Author        : JP
    Version       : 2.0
    Created       : 17/07/26
    Last Modified : 16/11/26
    Log Path      : C:\Logs\DesktopBackground\
#>
param(
    [switch]$DebugMode = $false
)

# Add necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to locate or automatically download IntuneWinAppUtil.exe
function Get-IntuneWinAppUtilPath {
    param([System.Windows.Forms.Label]$StatusLabel = $null)

    $candidatePaths = @()

    # 1. Check script directory or tools subfolder (local clone / batch launcher)
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidatePaths += (Join-Path $PSScriptRoot "IntuneWinAppUtil.exe")
        $candidatePaths += (Join-Path $PSScriptRoot "tools\IntuneWinAppUtil.exe")
    }

    # 2. Check user's local application data and persistent tool cache
    $localAppDataDir = Join-Path $env:LOCALAPPDATA "Microsoft\IntuneWinAppUtil"
    $customAppDataDir = Join-Path $env:LOCALAPPDATA "IntuneDesktopBackground"
    $printerAppDataDir = Join-Path $env:LOCALAPPDATA "IntunePrinterPackager"
    $candidatePaths += (Join-Path $localAppDataDir "IntuneWinAppUtil.exe")
    $candidatePaths += (Join-Path $customAppDataDir "IntuneWinAppUtil.exe")
    $candidatePaths += (Join-Path $printerAppDataDir "IntuneWinAppUtil.exe")
    $candidatePaths += (Join-Path $env:TEMP "IntuneWinAppUtil.exe")

    # 3. Check user downloads
    $userDownloads = Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) "Downloads"
    $candidatePaths += (Join-Path $userDownloads "IntuneWinAppUtil.exe")

    # 4. Check known sibling repository directories if in dev environment
    $scriptDir = $PSScriptRoot
    if (-not [string]::IsNullOrWhiteSpace($scriptDir)) {
        $intuneRoot = Split-Path -Parent $scriptDir
        if (-not [string]::IsNullOrWhiteSpace($intuneRoot)) {
            $candidatePaths += (Join-Path $intuneRoot "AddPrinter\IntuneWinAppUtil.exe")
            $candidatePaths += (Join-Path $intuneRoot "DesktopOverlay\IntuneWinAppUtil.exe")
            $candidatePaths += (Join-Path $intuneRoot "TeamsBackground\tools\IntuneWinAppUtil.exe")
        }
    }

    # 5. Check system PATH
    $cmd = try { Get-Command "IntuneWinAppUtil.exe" -ErrorAction Stop } catch { $null }
    if ($cmd -and [string]::IsNullOrWhiteSpace($cmd.Source) -eq $false -and (Test-Path -Path $cmd.Source -PathType Leaf)) {
        return $cmd.Source
    }

    # 6. Check existing candidate paths
    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -Path $cand -PathType Leaf)) {
            try {
                if ((Get-Item $cand).Length -gt 10000) {
                    return $cand
                }
            } catch {}
        }
    }

    # 7. Tool not found locally; determine persistent target download directory
    $downloadTargetDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
        $PSScriptRoot
    } else {
        $customAppDataDir
    }

    if (-not (Test-Path -Path $downloadTargetDir)) {
        try { New-Item -ItemType Directory -Path $downloadTargetDir -Force -ErrorAction Stop | Out-Null } catch {}
    }
    $targetExePath = Join-Path $downloadTargetDir "IntuneWinAppUtil.exe"

    # Prompt user to confirm auto-download
    $msg = "Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) was not found.`n`nWould you like to automatically download it from Microsoft's official GitHub repository?"
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Packer Tool Missing", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        if ($StatusLabel) {
            $StatusLabel.Text = "Downloading IntuneWinAppUtil.exe..."
            $StatusLabel.Refresh()
        }
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $downloadUrl = "https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe"

            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
            Invoke-WebRequest -Uri $downloadUrl -OutFile $targetExePath -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            $sig = Get-AuthenticodeSignature -FilePath $targetExePath -ErrorAction Stop
            if ($sig.Status -ne 'Valid') {
                if (Test-Path -LiteralPath $targetExePath) { try { Remove-Item $targetExePath -Force -ErrorAction Stop } catch {} }
                throw "Authenticode signature check failed for IntuneWinAppUtil.exe ($($sig.Status))"
            }
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default

            if (Test-Path -Path $targetExePath -PathType Leaf) {
                return $targetExePath
            } else {
                throw "Downloaded file was not found on disk."
            }
        } catch {
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Failed to download IntuneWinAppUtil.exe:`n`n$_`n`nPlease download it manually and place it in: $downloadTargetDir", "Download Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return $null
        }
    } else {
        return $null
    }
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

# Palette Colors (Obsidian Dark)
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

function New-BorderedSurface {
    param(
        [int]$PaddingAll = 14
    )
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

function Set-SecondaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.BackColor = $clrSecBtn
    $Button.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $clrBorder
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-UiFont -Size 9 -Style ([System.Drawing.FontStyle]::Bold)
}

function Set-PrimaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.BackColor = $clrTeal
    $Button.ForeColor = $clrTealText
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-UiFont -Size 9.5 -Style ([System.Drawing.FontStyle]::Bold)
}

# --- UI Configuration ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Desktop Background Packager v2.0"
Initialize-StandardWinForm -Form $form

# Fonts
$titleFont   = New-UiFont -Size 14.5 -Style ([System.Drawing.FontStyle]::Bold)
$sectionFont = New-UiFont -Size 9.5 -Style ([System.Drawing.FontStyle]::Bold)
$labelFont   = New-UiFont -Size 9.5
$textFont    = New-UiFont -Size 9
$monoFont    = New-UiFont -Size 8.5 -Monospace

# 1. Header Bar
$headerBar = New-Object System.Windows.Forms.Panel
$headerBar.Dock = [System.Windows.Forms.DockStyle]::Top
$headerBar.Height = 64
$headerBar.BackColor = $clrChrome
$headerBar.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 10)
$form.Controls.Add($headerBar)

$headerBorder = New-Object System.Windows.Forms.Panel
$headerBorder.Dock = [System.Windows.Forms.DockStyle]::Bottom
$headerBorder.Height = 1
$headerBorder.BackColor = $clrBorder
$headerBar.Controls.Add($headerBorder)

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
[void]$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$headerBar.Controls.Add($headerLayout)

$titleStack = New-Object System.Windows.Forms.Panel
$titleStack.Dock = [System.Windows.Forms.DockStyle]::Fill
$headerLayout.Controls.Add($titleStack, 0, 0)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Intune Desktop Background Packager"
$lblTitle.Font = $titleFont
$lblTitle.Location = New-Object System.Drawing.Point(0, 0)
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = $clrTextPrim
$titleStack.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Configure corporate desktop wallpaper and lock screen images into a Win32 .intunewin bundle"
$lblSubtitle.Font = $textFont
$lblSubtitle.Location = New-Object System.Drawing.Point(0, 24)
$lblSubtitle.AutoSize = $true
$lblSubtitle.ForeColor = $clrTextMuted
$titleStack.Controls.Add($lblSubtitle)

$versionWrap = New-Object System.Windows.Forms.FlowLayoutPanel
$versionWrap.AutoSize = $true
$versionWrap.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$versionWrap.WrapContents = $false
$versionWrap.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$headerLayout.Controls.Add($versionWrap, 1, 0)

$lblVersion = New-Object System.Windows.Forms.Label
$lblVersion.Text = "Package Version:"
$lblVersion.Font = $labelFont
$lblVersion.ForeColor = $clrTextMuted
$lblVersion.AutoSize = $true
$lblVersion.Margin = New-Object System.Windows.Forms.Padding(0, 5, 8, 0)
$versionWrap.Controls.Add($lblVersion)

$txtVersion = New-Object System.Windows.Forms.TextBox
$txtVersion.Text = "2.0"
$txtVersion.Font = $textFont
$txtVersion.BackColor = $clrInput
$txtVersion.ForeColor = $clrTextPrim
$txtVersion.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtVersion.Size = New-Object System.Drawing.Size(76, 24)
$versionWrap.Controls.Add($txtVersion)

# 2. Single-Row Footer Bar
$footerBar = New-Object System.Windows.Forms.Panel
$footerBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
$footerBar.Height = 58
$footerBar.BackColor = $clrChrome
$form.Controls.Add($footerBar)

$footerBorder = New-Object System.Windows.Forms.Panel
$footerBorder.Dock = [System.Windows.Forms.DockStyle]::Top
$footerBorder.Height = 1
$footerBorder.BackColor = $clrBorder
$footerBar.Controls.Add($footerBorder)

$footerTable = New-Object System.Windows.Forms.TableLayoutPanel
$footerTable.Dock = [System.Windows.Forms.DockStyle]::Fill
$footerTable.Padding = New-Object System.Windows.Forms.Padding(20, 10, 20, 10)
$footerTable.ColumnCount = 2
$footerTable.RowCount = 1
[void]$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$footerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$footerBar.Controls.Add($footerTable)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblStatus.AutoEllipsis = $true
$lblStatus.Font = $textFont
$lblStatus.ForeColor = $clrTextMuted
$lblStatus.Text = "Ready. Configure Desktop Background or Lock Screen options and click Generate."
$footerTable.Controls.Add($lblStatus, 0, 0)

$footerActions = New-Object System.Windows.Forms.FlowLayoutPanel
$footerActions.AutoSize = $true
$footerActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$footerActions.WrapContents = $false
$footerActions.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$footerTable.Controls.Add($footerActions, 1, 0)

$btnOpenOutput = New-Object System.Windows.Forms.Button
$btnOpenOutput.Text = "Open Output Folder"
$btnOpenOutput.Size = New-Object System.Drawing.Size(150, 34)
$btnOpenOutput.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-SecondaryButtonStyle $btnOpenOutput
$footerActions.Controls.Add($btnOpenOutput)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Intune Package"
$btnGenerate.Size = New-Object System.Drawing.Size(200, 34)
$btnGenerate.Margin = New-Object System.Windows.Forms.Padding(0)
Set-PrimaryButtonStyle $btnGenerate
$footerActions.Controls.Add($btnGenerate)

# 3. Main 2-Column Split Workspace
$mainSplit = New-Object System.Windows.Forms.TableLayoutPanel
$mainSplit.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainSplit.BackColor = $clrCanvas
$mainSplit.Padding = New-Object System.Windows.Forms.Padding(18, 14, 18, 14)
$mainSplit.ColumnCount = 2
$mainSplit.RowCount = 1
[void]$mainSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 58)))
[void]$mainSplit.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 42)))
[void]$mainSplit.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($mainSplit)
$mainSplit.BringToFront()

# Left Column Stack (Desktop + Lock Screen)
$leftCol = New-Object System.Windows.Forms.TableLayoutPanel
$leftCol.Dock = [System.Windows.Forms.DockStyle]::Fill
$leftCol.ColumnCount = 1
$leftCol.RowCount = 2
$leftCol.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
[void]$leftCol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 54)))
[void]$leftCol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 46)))
$mainSplit.Controls.Add($leftCol, 0, 0)

# Card 1: Desktop Background
$srcCard = New-BorderedSurface -PaddingAll 14
$srcCard.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$srcCard.Outer.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$leftCol.Controls.Add($srcCard.Outer, 0, 0)
$grpSource = $srcCard.Inner
$grpSource.AutoScroll = $true

$lblDesktopHead = New-Object System.Windows.Forms.Label
$lblDesktopHead.Text = "DESKTOP BACKGROUND"
$lblDesktopHead.Font = $sectionFont
$lblDesktopHead.ForeColor = $clrTeal
$lblDesktopHead.Location = New-Object System.Drawing.Point(12, 10)
$lblDesktopHead.AutoSize = $true
$grpSource.Controls.Add($lblDesktopHead)

$chkSetDesktop = New-Object System.Windows.Forms.CheckBox
$chkSetDesktop.Text = "Enable Desktop Background Management"
$chkSetDesktop.Checked = $true
$chkSetDesktop.ForeColor = $clrTextPrim
$chkSetDesktop.Location = New-Object System.Drawing.Point(14, 34)
$chkSetDesktop.AutoSize = $true
$grpSource.Controls.Add($chkSetDesktop)

$radLocal = New-Object System.Windows.Forms.RadioButton
$radLocal.Text = "Local Image File"
$radLocal.Location = New-Object System.Drawing.Point(14, 66)
$radLocal.Checked = $true
$radLocal.ForeColor = $clrTextSec
$radLocal.AutoSize = $true
$grpSource.Controls.Add($radLocal)

$radUrl = New-Object System.Windows.Forms.RadioButton
$radUrl.Text = "Image URL (Dynamic)"
$radUrl.Location = New-Object System.Drawing.Point(14, 100)
$radUrl.ForeColor = $clrTextSec
$radUrl.AutoSize = $true
$grpSource.Controls.Add($radUrl)

$txtLocalFile = New-Object System.Windows.Forms.TextBox
$txtLocalFile.Location = New-Object System.Drawing.Point(185, 64)
$txtLocalFile.Size = New-Object System.Drawing.Size(210, 24)
$txtLocalFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtLocalFile.BackColor = $clrInput
$txtLocalFile.ForeColor = $clrTextPrim
$txtLocalFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLocalFile.Font = $textFont
$grpSource.Controls.Add($txtLocalFile)

$btnBrowseLocal = New-Object System.Windows.Forms.Button
$btnBrowseLocal.Text = "Browse..."
$btnBrowseLocal.Location = New-Object System.Drawing.Point(403, 62)
$btnBrowseLocal.Size = New-Object System.Drawing.Size(86, 27)
$btnBrowseLocal.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Set-SecondaryButtonStyle $btnBrowseLocal
$grpSource.Controls.Add($btnBrowseLocal)

$txtUrl = New-Object System.Windows.Forms.TextBox
$txtUrl.Location = New-Object System.Drawing.Point(185, 98)
$txtUrl.Size = New-Object System.Drawing.Size(304, 24)
$txtUrl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtUrl.BackColor = $clrInput
$txtUrl.ForeColor = $clrTextPrim
$txtUrl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtUrl.Font = $textFont
$txtUrl.Enabled = $false
$grpSource.Controls.Add($txtUrl)

$lblWallpaperStyle = New-Object System.Windows.Forms.Label
$lblWallpaperStyle.Text = "Wallpaper Style:"
$lblWallpaperStyle.Location = New-Object System.Drawing.Point(14, 136)
$lblWallpaperStyle.ForeColor = $clrTextSec
$lblWallpaperStyle.AutoSize = $true
$lblWallpaperStyle.Font = $textFont
$grpSource.Controls.Add($lblWallpaperStyle)

$cmbWallpaperStyle = New-Object System.Windows.Forms.ComboBox
$cmbWallpaperStyle.Location = New-Object System.Drawing.Point(185, 132)
$cmbWallpaperStyle.Size = New-Object System.Drawing.Size(160, 25)
$cmbWallpaperStyle.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbWallpaperStyle.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cmbWallpaperStyle.BackColor = $clrInput
$cmbWallpaperStyle.ForeColor = $clrTextPrim
$cmbWallpaperStyle.Font = $textFont
[void]$cmbWallpaperStyle.Items.AddRange(@("Fill", "Fit", "Stretch", "Center", "Span", "Tile"))
$cmbWallpaperStyle.SelectedIndex = 0
$grpSource.Controls.Add($cmbWallpaperStyle)

$chkLockBackground = New-Object System.Windows.Forms.CheckBox
$chkLockBackground.Text = "Prevent users from changing the Desktop Background"
$chkLockBackground.ForeColor = $clrTextSec
$chkLockBackground.Location = New-Object System.Drawing.Point(14, 170)
$chkLockBackground.AutoSize = $true
$grpSource.Controls.Add($chkLockBackground)

# Card 2: Lock Screen
$lockCard = New-BorderedSurface -PaddingAll 14
$lockCard.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$leftCol.Controls.Add($lockCard.Outer, 0, 1)
$grpLockScreen = $lockCard.Inner
$grpLockScreen.AutoScroll = $true

$lblLockHead = New-Object System.Windows.Forms.Label
$lblLockHead.Text = "LOCK SCREEN"
$lblLockHead.Font = $sectionFont
$lblLockHead.ForeColor = $clrTeal
$lblLockHead.Location = New-Object System.Drawing.Point(12, 10)
$lblLockHead.AutoSize = $true
$grpLockScreen.Controls.Add($lblLockHead)

$chkSetLockScreen = New-Object System.Windows.Forms.CheckBox
$chkSetLockScreen.Text = "Enable Lock Screen Management"
$chkSetLockScreen.Checked = $true
$chkSetLockScreen.ForeColor = $clrTextPrim
$chkSetLockScreen.Location = New-Object System.Drawing.Point(14, 34)
$chkSetLockScreen.AutoSize = $true
$grpLockScreen.Controls.Add($chkSetLockScreen)

$radLockSame = New-Object System.Windows.Forms.RadioButton
$radLockSame.Text = "Same as Desktop Background"
$radLockSame.Location = New-Object System.Drawing.Point(14, 62)
$radLockSame.Checked = $true
$radLockSame.ForeColor = $clrTextSec
$radLockSame.AutoSize = $true
$grpLockScreen.Controls.Add($radLockSame)

$radLockLocal = New-Object System.Windows.Forms.RadioButton
$radLockLocal.Text = "Local Image File"
$radLockLocal.Location = New-Object System.Drawing.Point(14, 92)
$radLockLocal.ForeColor = $clrTextSec
$radLockLocal.AutoSize = $true
$grpLockScreen.Controls.Add($radLockLocal)

$radLockUrl = New-Object System.Windows.Forms.RadioButton
$radLockUrl.Text = "Image URL (Dynamic)"
$radLockUrl.Location = New-Object System.Drawing.Point(14, 124)
$radLockUrl.ForeColor = $clrTextSec
$radLockUrl.AutoSize = $true
$grpLockScreen.Controls.Add($radLockUrl)

$txtLockLocalFile = New-Object System.Windows.Forms.TextBox
$txtLockLocalFile.Location = New-Object System.Drawing.Point(185, 90)
$txtLockLocalFile.Size = New-Object System.Drawing.Size(210, 24)
$txtLockLocalFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtLockLocalFile.BackColor = $clrInput
$txtLockLocalFile.ForeColor = $clrTextPrim
$txtLockLocalFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLockLocalFile.Font = $textFont
$txtLockLocalFile.Enabled = $false
$grpLockScreen.Controls.Add($txtLockLocalFile)

$btnBrowseLockLocal = New-Object System.Windows.Forms.Button
$btnBrowseLockLocal.Text = "Browse..."
$btnBrowseLockLocal.Location = New-Object System.Drawing.Point(403, 88)
$btnBrowseLockLocal.Size = New-Object System.Drawing.Size(86, 27)
$btnBrowseLockLocal.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Set-SecondaryButtonStyle $btnBrowseLockLocal
$btnBrowseLockLocal.Enabled = $false
$grpLockScreen.Controls.Add($btnBrowseLockLocal)

$txtLockUrl = New-Object System.Windows.Forms.TextBox
$txtLockUrl.Location = New-Object System.Drawing.Point(185, 122)
$txtLockUrl.Size = New-Object System.Drawing.Size(304, 24)
$txtLockUrl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtLockUrl.BackColor = $clrInput
$txtLockUrl.ForeColor = $clrTextPrim
$txtLockUrl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtLockUrl.Font = $textFont
$txtLockUrl.Enabled = $false
$grpLockScreen.Controls.Add($txtLockUrl)

$chkLockLockScreen = New-Object System.Windows.Forms.CheckBox
$chkLockLockScreen.Text = "Prevent users from changing the Lock Screen"
$chkLockLockScreen.ForeColor = $clrTextSec
$chkLockLockScreen.Location = New-Object System.Drawing.Point(14, 156)
$chkLockLockScreen.AutoSize = $true
$grpLockScreen.Controls.Add($chkLockLockScreen)

# Right Column Stack (Output + Preview & Bundle Artifacts)
$rightCol = New-Object System.Windows.Forms.TableLayoutPanel
$rightCol.Dock = [System.Windows.Forms.DockStyle]::Fill
$rightCol.ColumnCount = 1
$rightCol.RowCount = 2
$rightCol.Margin = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
[void]$rightCol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 102)))
[void]$rightCol.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainSplit.Controls.Add($rightCol, 1, 0)

# Output Directory Card
$outCard = New-BorderedSurface -PaddingAll 14
$outCard.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$outCard.Outer.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$rightCol.Controls.Add($outCard.Outer, 0, 0)
$grpOutput = $outCard.Inner

$lblOutHead = New-Object System.Windows.Forms.Label
$lblOutHead.Text = "OUTPUT DIRECTORY"
$lblOutHead.Font = $sectionFont
$lblOutHead.ForeColor = $clrTeal
$lblOutHead.Location = New-Object System.Drawing.Point(12, 10)
$lblOutHead.AutoSize = $true
$grpOutput.Controls.Add($lblOutHead)

$txtOutputDir = New-Object System.Windows.Forms.TextBox
$txtOutputDir.Location = New-Object System.Drawing.Point(14, 38)
$txtOutputDir.Size = New-Object System.Drawing.Size(220, 24)
$txtOutputDir.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$txtOutputDir.BackColor = $clrInput
$txtOutputDir.ForeColor = $clrTextPrim
$txtOutputDir.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtOutputDir.Font = $textFont

$defaultOutput = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
    Join-Path $PSScriptRoot "Output"
} else {
    Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) "Downloads\DesktopBackground_Output"
}
$txtOutputDir.Text = $defaultOutput
$grpOutput.Controls.Add($txtOutputDir)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = "Browse..."
$btnBrowseOutput.Location = New-Object System.Drawing.Point(242, 36)
$btnBrowseOutput.Size = New-Object System.Drawing.Size(86, 27)
$btnBrowseOutput.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Set-SecondaryButtonStyle $btnBrowseOutput
$grpOutput.Controls.Add($btnBrowseOutput)

# Preview & Bundle Artifacts Card
$previewCard = New-BorderedSurface -PaddingAll 14
$previewCard.Outer.Dock = [System.Windows.Forms.DockStyle]::Fill
$rightCol.Controls.Add($previewCard.Outer, 0, 1)
$grpPreview = $previewCard.Inner

$previewLayout = New-Object System.Windows.Forms.TableLayoutPanel
$previewLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewLayout.ColumnCount = 1
$previewLayout.RowCount = 3
[void]$previewLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$previewLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 52)))
[void]$previewLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 48)))
$grpPreview.Controls.Add($previewLayout)

$lblPrevHead = New-Object System.Windows.Forms.Label
$lblPrevHead.Text = "IMAGE PREVIEW & BUNDLE ARTIFACTS"
$lblPrevHead.Font = $sectionFont
$lblPrevHead.ForeColor = $clrTeal
$lblPrevHead.AutoSize = $true
$lblPrevHead.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
$previewLayout.Controls.Add($lblPrevHead, 0, 0)

$picContainer = New-Object System.Windows.Forms.Panel
$picContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$picContainer.BackColor = $clrLog
$picContainer.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$previewLayout.Controls.Add($picContainer, 0, 1)

$picPreview = New-Object System.Windows.Forms.PictureBox
$picPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$picPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$picPreview.BackColor = $clrLog
$picContainer.Controls.Add($picPreview)

$lblPreviewHint = New-Object System.Windows.Forms.Label
$lblPreviewHint.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblPreviewHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblPreviewHint.ForeColor = $clrTextMuted
$lblPreviewHint.Font = $textFont
$lblPreviewHint.Text = "Select a local image file to preview"
$picContainer.Controls.Add($lblPreviewHint)
$lblPreviewHint.BringToFront()

$txtBundleLog = New-Object System.Windows.Forms.TextBox
$txtBundleLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtBundleLog.Multiline = $true
$txtBundleLog.ReadOnly = $true
$txtBundleLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtBundleLog.BackColor = $clrLog
$txtBundleLog.ForeColor = $clrTextSec
$txtBundleLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$txtBundleLog.Font = $monoFont
$txtBundleLog.Text = @(
    "GENERATED ARTIFACTS:",
    " - DesktopBackground_v<Ver>_DDMMYY.intunewin",
    " - DetectionScript.ps1",
    " - InstallCommands.txt (5-step Intune guide)",
    "",
    "WORKSTATION LOGS:",
    " C:\Logs\DesktopBackground\",
    "   Install.log | ApplyBackground.log",
    "   Detection.log | Uninstall.log"
) -join "`r`n"
$previewLayout.Controls.Add($txtBundleLog, 0, 2)

function Update-ImagePreview {
    param([string]$ImagePath)
    if ($null -ne $picPreview.Image) {
        $oldImg = $picPreview.Image
        $picPreview.Image = $null
        $oldImg.Dispose()
    }
    if (-not [string]::IsNullOrWhiteSpace($ImagePath) -and (Test-Path -LiteralPath $ImagePath -PathType Leaf)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
            $ms = New-Object System.IO.MemoryStream(, $bytes)
            $picPreview.Image = [System.Drawing.Image]::FromStream($ms)
            $lblPreviewHint.Visible = $false
            return
        } catch {}
    }
    $lblPreviewHint.Visible = $true
}

# Events for Source Type switching
$radLocal.Add_CheckedChanged({
    if ($chkSetDesktop.Checked) {
        $txtLocalFile.Enabled = $radLocal.Checked
        $btnBrowseLocal.Enabled = $radLocal.Checked
        $txtUrl.Enabled = !$radLocal.Checked
    }
})

$chkSetDesktop.Add_CheckedChanged({
    $enabled = $chkSetDesktop.Checked
    $radLocal.Enabled = $enabled
    $radUrl.Enabled = $enabled
    $lblWallpaperStyle.Enabled = $enabled
    $cmbWallpaperStyle.Enabled = $enabled
    $chkLockBackground.Enabled = $enabled

    if ($enabled) {
        $txtLocalFile.Enabled = $radLocal.Checked
        $btnBrowseLocal.Enabled = $radLocal.Checked
        $txtUrl.Enabled = !$radLocal.Checked
    } else {
        $txtLocalFile.Enabled = $false
        $btnBrowseLocal.Enabled = $false
        $txtUrl.Enabled = $false
    }

    if ($chkSetLockScreen.Checked) {
        if (!$enabled) {
            if ($radLockSame.Checked) {
                $radLockLocal.Checked = $true
            }
            $radLockSame.Enabled = $false
        } else {
            $radLockSame.Enabled = $true
        }
    }
})

$btnBrowseLocal.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLocalFile.Text = $openFileDialog.FileName
        Update-ImagePreview -ImagePath $openFileDialog.FileName
    }
})

# Lock Screen UI Events
$radLockLocal.Add_CheckedChanged({
    if ($chkSetLockScreen.Checked) {
        $txtLockLocalFile.Enabled = $radLockLocal.Checked
        $btnBrowseLockLocal.Enabled = $radLockLocal.Checked
    }
})

$radLockUrl.Add_CheckedChanged({
    if ($chkSetLockScreen.Checked) {
        $txtLockUrl.Enabled = $radLockUrl.Checked
    }
})

$chkSetLockScreen.Add_CheckedChanged({
    $enabled = $chkSetLockScreen.Checked
    $radLockSame.Enabled = $enabled -and $chkSetDesktop.Checked
    $radLockLocal.Enabled = $enabled
    $radLockUrl.Enabled = $enabled
    $chkLockLockScreen.Enabled = $enabled

    if ($enabled) {
        if (!$chkSetDesktop.Checked -and $radLockSame.Checked) {
            $radLockLocal.Checked = $true
        }
        $txtLockLocalFile.Enabled = $radLockLocal.Checked
        $btnBrowseLockLocal.Enabled = $radLockLocal.Checked
        $txtLockUrl.Enabled = $radLockUrl.Checked
    } else {
        $txtLockLocalFile.Enabled = $false
        $btnBrowseLockLocal.Enabled = $false
        $txtLockUrl.Enabled = $false
    }
})

$btnBrowseLockLocal.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All Files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLockLocalFile.Text = $openFileDialog.FileName
        if (-not $chkSetDesktop.Checked -or -not (Test-Path -LiteralPath $txtLocalFile.Text)) {
            Update-ImagePreview -ImagePath $openFileDialog.FileName
        }
    }
})

$btnBrowseOutput.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.SelectedPath = $txtOutputDir.Text
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $folderBrowser.SelectedPath
    }
})

$btnOpenOutput.Add_Click({
    $dir = $txtOutputDir.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        if (-not (Test-Path -Path $dir)) {
            try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch {}
        }
        if (Test-Path -Path $dir) {
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$dir`""
        }
    }
})

# Main Processing Logic
$btnGenerate.Add_Click({
    $lblStatus.Text = "Starting packaging process..."
    $lblStatus.Refresh()

    try {
        # 1. Validation
        $setDesktop = $chkSetDesktop.Checked
        $setLockScreen = $chkSetLockScreen.Checked
        $lockIsSame = $radLockSame.Checked
        $lockIsLocalDiff = $radLockLocal.Checked
        $lockIsUrlDiff = $radLockUrl.Checked
        $lockIsDiff = $lockIsLocalDiff -or $lockIsUrlDiff

        # Package Version
        $rawVersion = $txtVersion.Text.Trim()
        $cleanVersion = ($rawVersion -replace '[^a-zA-Z0-9\.\-]', '')
        if ([string]::IsNullOrWhiteSpace($cleanVersion)) { $cleanVersion = "2.0" }
        $versionTag = if ($cleanVersion -match '^[vV]') { $cleanVersion } else { "v$cleanVersion" }

        # Wallpaper Style Mapping
        $selectedStyle = if ($cmbWallpaperStyle.SelectedItem) { $cmbWallpaperStyle.SelectedItem.ToString() } else { "Fill" }
        switch ($selectedStyle) {
            "Fill"    { $wallpaperStyleVal = "10"; $tileWallpaperVal = "0" }
            "Fit"     { $wallpaperStyleVal = "6";  $tileWallpaperVal = "0" }
            "Stretch" { $wallpaperStyleVal = "2";  $tileWallpaperVal = "0" }
            "Center"  { $wallpaperStyleVal = "0";  $tileWallpaperVal = "0" }
            "Tile"    { $wallpaperStyleVal = "0";  $tileWallpaperVal = "1" }
            "Span"    { $wallpaperStyleVal = "22"; $tileWallpaperVal = "0" }
            default   { $wallpaperStyleVal = "10"; $tileWallpaperVal = "0" }
        }

        if (!$setDesktop -and !$setLockScreen) {
            [System.Windows.Forms.MessageBox]::Show("You must configure at least one option: Desktop Background or Lock Screen.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if ($setDesktop) {
            if ($radLocal.Checked -and (-not (Test-Path $txtLocalFile.Text))) {
                [System.Windows.Forms.MessageBox]::Show("Please select a valid local image file for the Desktop Background.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            if ($radUrl.Checked -and [string]::IsNullOrWhiteSpace($txtUrl.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid Desktop Background image URL.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }

        if ($setLockScreen -and $lockIsDiff) {
            if ($lockIsUrlDiff) {
                if ([string]::IsNullOrWhiteSpace($txtLockUrl.Text)) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter a valid URL for the Lock Screen.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            } else {
                if (-not (Test-Path $txtLockLocalFile.Text)) {
                    [System.Windows.Forms.MessageBox]::Show("Please select a valid local image file for the Lock Screen.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            }
        }

        $outDir = $txtOutputDir.Text.Trim()
        if (-not (Test-Path $outDir)) {
            try {
                New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Could not create output directory: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }

        # Determine if we are packaging local files vs URLs
        $isLocal = $true
        if ($setDesktop -and $radUrl.Checked) {
            $isLocal = $false
        }
        if ($setLockScreen -and $lockIsUrlDiff) {
            $isLocal = $false
        }

        # 2. Check/Download IntuneWinAppUtil.exe
        $packerExe = Get-IntuneWinAppUtilPath -StatusLabel $lblStatus
        if ([string]::IsNullOrWhiteSpace($packerExe) -or (-not (Test-Path -Path $packerExe -PathType Leaf))) {
            $lblStatus.Text = "Build aborted: IntuneWinAppUtil.exe is required."
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            return
        }

        # 3. Setup Staging Directory
        $lblStatus.Text = "Setting up staging directory..."
        $lblStatus.Refresh()
        $stagingBase = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
            $PSScriptRoot
        } else {
            $env:TEMP
        }
        $stagingDir = Join-Path $stagingBase "Staging_IntuneBackground_$([guid]::NewGuid().ToString().Substring(0,8))"
        if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

        # 4. Generate Install.ps1
        $installScriptPath = Join-Path $stagingDir "Install.ps1"
        $installContent = ""

        if ($isLocal) {
            # Copy Desktop Image to Staging if enabled
            $fileName = ""
            if ($setDesktop) {
                $fileName = Split-Path $txtLocalFile.Text -Leaf
                $destImg = Join-Path $stagingDir "desktop_$fileName"
                Copy-Item $txtLocalFile.Text -Destination $destImg -Force
            }

            $lockFileName = ""
            if ($setLockScreen) {
                if ($lockIsSame) {
                    $lockFileName = "desktop_$fileName"
                } else {
                    $lockOriginalName = Split-Path $txtLockLocalFile.Text -Leaf
                    $lockFileName = "lock_$lockOriginalName"
                    $destLockImg = Join-Path $stagingDir $lockFileName
                    Copy-Item $txtLockLocalFile.Text -Destination $destLockImg -Force
                }
            }

            # Create ApplyBackground.ps1 (Only if Desktop is enabled)
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
                Set-Content -Path (Join-Path $stagingDir "ApplyBackground.ps1") -Value $applyScriptContent -Force
            }

            # Create Install.ps1
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
                if ($chkLockBackground.Checked) {
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
                if ($chkLockLockScreen.Checked) {
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
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
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
            # URL logic
            $desktopUrl = if ($setDesktop -and $radUrl.Checked) { $txtUrl.Text } else { "" }
            $lockUrl = if ($setLockScreen) {
                if ($lockIsSame) { $desktopUrl } elseif ($lockIsUrlDiff) { $txtLockUrl.Text }
            } else { "" }
            
            $desktopExt = ".png"
            if ($setDesktop -and $radUrl.Checked) {
                $desktopExt = [System.IO.Path]::GetExtension($desktopUrl)
                if ($desktopExt -notmatch '\.(jpg|jpeg|png|bmp)') { $desktopExt = '.png' }
            }
            
            $lockExt = ".png"
            if ($setLockScreen) {
                if ($lockIsSame) {
                    $lockExt = $desktopExt
                } elseif ($lockIsUrlDiff) {
                    $lockExt = [System.IO.Path]::GetExtension($lockUrl)
                    if ($lockExt -notmatch '\.(jpg|jpeg|png|bmp)') { $lockExt = '.png' }
                }
            }

            $fileName = ""
            if ($setDesktop -and $radLocal.Checked) {
                $fileName = Split-Path $txtLocalFile.Text -Leaf
                $destImg = Join-Path $stagingDir "desktop_$fileName"
                Copy-Item $txtLocalFile.Text -Destination $destImg -Force
            }

            $lockFileName = ""
            if ($setLockScreen -and $lockIsLocalDiff) {
                $lockOriginalName = Split-Path $txtLockLocalFile.Text -Leaf
                $lockFileName = "lock_$lockOriginalName"
                $destLockImg = Join-Path $stagingDir $lockFileName
                Copy-Item $txtLockLocalFile.Text -Destination $destLockImg -Force
            }

            $destDesktopPath = ""
            if ($setDesktop) {
                if ($radLocal.Checked) {
                    $destDesktopPath = "`$targetPath\desktop_$fileName"
                } else {
                    $destDesktopPath = "`$targetPath\desktop_bg$desktopExt"
                }
            }

            $destLockPath = ""
            if ($setLockScreen) {
                if ($lockIsSame) {
                    if ($radLocal.Checked) {
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
            
            # Create DownloadBackgrounds.ps1
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

            if ($setDesktop -and $radUrl.Checked) {
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
                } elseif ($lockIsSame -and $radUrl.Checked) {
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
            
            if ($setDesktop -and $chkLockBackground.Checked) {
                $dlContent += @"

try { Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction Stop } catch {}
"@
            }

            if ($setDesktop) {
                $dlContent += @"

try { Start-ScheduledTask -TaskName 'ApplyCorporateBackground' -ErrorAction Stop } catch {}
"@
            }

            if ($setDesktop -and $chkLockBackground.Checked) {
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
            Set-Content -Path (Join-Path $stagingDir "DownloadBackgrounds.ps1") -Value $dlContent -Force

            # Create ApplyBackground.ps1 (Only if Desktop is enabled)
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
                Set-Content -Path (Join-Path $stagingDir "ApplyBackground.ps1") -Value $applyScriptContent -Force
            }

            # Create Install.ps1
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
                if ($radLocal.Checked) {
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
                if ($chkLockBackground.Checked) {
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
                if ($chkLockLockScreen.Checked) {
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
            try { Remove-ItemProperty -Path `"Registry::HKEY_USERS\`$sid\Software\Policies\Microsoft\Windows\Personalization`" -Name `"NoChangingLockScreen`" -ErrorAction Stop | Out-Null } catch {}
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
        Set-Content -Path (Join-Path $stagingDir "Uninstall.ps1") -Value $uninstallContent -Force

        # Trim output directory to prevent trailing spaces
        $outDir = $txtOutputDir.Text.Trim()
        
        # Ensure output directory exists
        if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        
        $process = Start-Process -FilePath $packerExe -ArgumentList "-c `"$stagingDir`"", "-s `"Install.ps1`"", "-o `"$outDir`"", "-q" -NoNewWindow -PassThru -Wait
        
        if ($process.ExitCode -ne 0) {
            throw "IntuneWinAppUtil failed to package the application. Exit code: $($process.ExitCode)"
        }

        # Rename output .intunewin file to DesktopBackground_v<version>_DDMMYY.intunewin
        $dateStr = Get-Date -Format "ddMMyy"
        $newWimName = "DesktopBackground_${versionTag}_$dateStr.intunewin"
        $wimPath = Join-Path $outDir "Install.intunewin"
        $newWimPath = Join-Path $outDir $newWimName
        
        # Give the filesystem a moment to index the new file
        Start-Sleep -Milliseconds 500
        
        if (Test-Path -LiteralPath $wimPath) {
            if (Test-Path -LiteralPath $newWimPath) { try { Remove-Item -LiteralPath $newWimPath -Force -ErrorAction Stop } catch {} }
            
            # Retry loop to handle OneDrive file locks
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
            if (!$renamed) {
                throw "Failed to rename output package (file is locked by another process like OneDrive). You can manually rename 'Install.intunewin' to '$newWimName'."
            }
        } else {
            throw "Packaging failed: Could not find 'Install.intunewin' at the expected path: '$wimPath'."
        }

        # 6. Generate Outputs (Detection and Install Commands)
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
                # Lock Screen Only (Local)
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
            # URL (Desktop or Lock Screen or Both)
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
        $detectionFile = Join-Path $outDir "DetectionScript.ps1"
        Set-Content -Path $detectionFile -Value $detectionContent -Force

        $generatedTimestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
        $createdDateStr = Get-Date -Format "dd/MM/yy"
        $desktopSummary = if ($setDesktop) { "Enabled (Style: $selectedStyle | Lock: $($chkLockBackground.Checked))" } else { "Disabled" }
        $lockSummary    = if ($setLockScreen) { "Enabled (Lock: $($chkLockLockScreen.Checked))" } else { "Disabled" }

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
        $commandsFile = Join-Path $outDir "InstallCommands.txt"
        Set-Content -Path $commandsFile -Value $commandsContent -Force

        # 7. Cleanup
        if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }

        $lblStatus.Text = "Package successfully created in Output Directory!"
        $lblStatus.ForeColor = $clrOk
        $txtBundleLog.AppendText("`r`n[SUCCESS] Created $newWimName, DetectionScript.ps1, and InstallCommands.txt in $outDir")
        [System.Windows.Forms.MessageBox]::Show("Packaging Complete! Check the output directory for your $newWimName, DetectionScript.ps1, and InstallCommands.txt.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        if ($stagingDir -and (Test-Path $stagingDir)) { try { Remove-Item $stagingDir -Recurse -Force -ErrorAction Stop } catch {} }
        $lblStatus.Text = "Error: $($_.Exception.Message)"
        $lblStatus.ForeColor = $clrErr
        $txtBundleLog.AppendText("`r`n[ERROR] $($_.Exception.Message)")
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Show Form
$form.ShowDialog() | Out-Null
$form.Dispose()
