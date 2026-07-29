<#
.SYNOPSIS
    IT Support Overlay Studio - Interactive Design & Package Builder
.DESCRIPTION
    A visual Windows Forms application allowing IT administrators to design,
    customize, preview, test, and package the IT Support Desktop Overlay into an Intune Win32 (.intunewin) file.
#>

# --- In-Memory Guard for Bundled Packages ---
if ([string]::IsNullOrEmpty($PSScriptRoot) -or $PSScriptRoot -match "^iex") {
    throw "In-memory execution (via iex / irm) is not supported for bundled packages. OverlayStudio.ps1 relies on local relative files (src/ payload directory and IntuneWinAppUtil.exe). Please download and extract the repository package locally before executing."
}

# --- Resolve Absolute Script Root Folder ---
$Script:RootPath = $PSScriptRoot
if ([string]::IsNullOrEmpty($Script:RootPath)) {
    $Script:RootPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(".")
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
    try { Add-Type -TypeDefinition $DpiSignature -ErrorAction SilentlyContinue } catch {}
}
try { [StudioDpiHelper]::SetProcessDPIAware() | Out-Null } catch {}

# --- Config State ---
$Config = [ordered]@{
    Title        = "IT SUPPORT HELPDESK"
    Phone        = "0800 123 4567"
    Email        = "support@company.com"
    Position     = "BottomRight"
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

# --- Form Setup ---
$studio = New-Object System.Windows.Forms.Form
$studio.Text            = "IT Support Overlay Studio and Package Builder"
$studio.Size            = New-Object System.Drawing.Size(920, 680)
$studio.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$studio.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$studio.MaximizeBox     = $false
$studio.BackColor       = [System.Drawing.Color]::FromArgb(15, 23, 42)
$studio.ForeColor       = [System.Drawing.Color]::White
$studio.Font            = New-Object System.Drawing.Font("Segoe UI", 9.5)

# --- Header ---
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock      = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height    = 60
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$studio.Controls.Add($headerPanel)

$headerTitle = New-Object System.Windows.Forms.Label
$headerTitle.Text      = "IT Support Overlay Studio"
$headerTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$headerTitle.ForeColor = [System.Drawing.Color]::FromArgb(56, 189, 248)
$headerTitle.Location  = New-Object System.Drawing.Point(16, 14)
$headerTitle.AutoSize  = $true
$headerPanel.Controls.Add($headerTitle)

$headerSub = New-Object System.Windows.Forms.Label
$headerSub.Text      = "Customize design, preview live, and export Intune (.intunewin) package"
$headerSub.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5)
$headerSub.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
$headerSub.Location  = New-Object System.Drawing.Point(340, 20)
$headerSub.AutoSize  = $true
$headerPanel.Controls.Add($headerSub)

# Left Column (Settings Controls)
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Location  = New-Object System.Drawing.Point(16, 76)
$leftPanel.Size      = New-Object System.Drawing.Size(445, 540)
$leftPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$leftPanel.Padding   = New-Object System.Windows.Forms.Padding(12)
$studio.Controls.Add($leftPanel)

# Right Column (Live Preview & Export Actions)
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Location  = New-Object System.Drawing.Point(477, 76)
$rightPanel.Size      = New-Object System.Drawing.Size(410, 540)
$rightPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$rightPanel.Padding   = New-Object System.Windows.Forms.Padding(12)
$studio.Controls.Add($rightPanel)

$leftPanel.BringToFront()
$rightPanel.BringToFront()

function Add-SectionLabel ($parent, $text, $yPos) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $text
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(56, 189, 248)
    $lbl.Location  = New-Object System.Drawing.Point(12, $yPos)
    $lbl.AutoSize  = $true
    $parent.Controls.Add($lbl)
}

# --- LEFT PANEL: CONTROLS ---

# Section 1: Content & Branding
Add-SectionLabel $leftPanel "1. Content and Branding" 12

# Title Text
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Header Title:"; $lblTitle.Location = New-Object System.Drawing.Point(12, 42); $lblTitle.AutoSize = $true; $leftPanel.Controls.Add($lblTitle)
$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Text = $Config.Title; $txtTitle.Location = New-Object System.Drawing.Point(130, 39); $txtTitle.Size = New-Object System.Drawing.Size(295, 26); $leftPanel.Controls.Add($txtTitle)

# Phone
$lblPhone = New-Object System.Windows.Forms.Label
$lblPhone.Text = "Support Phone:"; $lblPhone.Location = New-Object System.Drawing.Point(12, 76); $lblPhone.AutoSize = $true; $leftPanel.Controls.Add($lblPhone)
$txtPhone = New-Object System.Windows.Forms.TextBox
$txtPhone.Text = $Config.Phone; $txtPhone.Location = New-Object System.Drawing.Point(130, 73); $txtPhone.Size = New-Object System.Drawing.Size(295, 26); $leftPanel.Controls.Add($txtPhone)

# Email
$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "Support Email:"; $lblEmail.Location = New-Object System.Drawing.Point(12, 110); $lblEmail.AutoSize = $true; $leftPanel.Controls.Add($lblEmail)
$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Text = $Config.Email; $txtEmail.Location = New-Object System.Drawing.Point(130, 107); $txtEmail.Size = New-Object System.Drawing.Size(295, 26); $leftPanel.Controls.Add($txtEmail)

# Section 2: Colors & Theme
Add-SectionLabel $leftPanel "2. Theme Colors and Opacity" 146

# Accent Color Picker Button
$btnAccent = New-Object System.Windows.Forms.Button
$btnAccent.Text      = "Accent Color"
$btnAccent.Location  = New-Object System.Drawing.Point(12, 175)
$btnAccent.Size      = New-Object System.Drawing.Size(125, 30)
$btnAccent.BackColor = $Config.AccentColor
$btnAccent.ForeColor = [System.Drawing.Color]::White
$btnAccent.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$leftPanel.Controls.Add($btnAccent)

# Background Color Picker Button
$btnBg = New-Object System.Windows.Forms.Button
$btnBg.Text      = "Background"
$btnBg.Location  = New-Object System.Drawing.Point(150, 175)
$btnBg.Size      = New-Object System.Drawing.Size(125, 30)
$btnBg.BackColor = $Config.BgColor
$btnBg.ForeColor = [System.Drawing.Color]::White
$btnBg.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$leftPanel.Controls.Add($btnBg)

# Text Color Picker Button
$btnText = New-Object System.Windows.Forms.Button
$btnText.Text      = "Text Color"
$btnText.Location  = New-Object System.Drawing.Point(288, 175)
$btnText.Size      = New-Object System.Drawing.Size(137, 30)
$btnText.BackColor = $Config.TextColor
$btnText.ForeColor = [System.Drawing.Color]::Black
$btnText.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$leftPanel.Controls.Add($btnText)

# Opacity Slider
$lblOpacity = New-Object System.Windows.Forms.Label
$lblOpacity.Text = "Card Opacity (92%):"; $lblOpacity.Location = New-Object System.Drawing.Point(12, 218); $lblOpacity.AutoSize = $true; $leftPanel.Controls.Add($lblOpacity)
$trackOpacity = New-Object System.Windows.Forms.TrackBar
$trackOpacity.Minimum = 60; $trackOpacity.Maximum = 100; $trackOpacity.Value = 92
$trackOpacity.Location = New-Object System.Drawing.Point(150, 212); $trackOpacity.Size = New-Object System.Drawing.Size(275, 45)
$leftPanel.Controls.Add($trackOpacity)

# Section 3: Position & Layer
Add-SectionLabel $leftPanel "3. Placement and Z-Order Layer" 265

# Position Combo
$lblPos = New-Object System.Windows.Forms.Label
$lblPos.Text = "Screen Corner:"; $lblPos.Location = New-Object System.Drawing.Point(12, 298); $lblPos.AutoSize = $true; $leftPanel.Controls.Add($lblPos)
$comboPos = New-Object System.Windows.Forms.ComboBox
$comboPos.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboPos.Items.AddRange(@("BottomRight", "TopRight", "BottomLeft", "TopLeft"))
$comboPos.SelectedItem = "BottomRight"
$comboPos.Location = New-Object System.Drawing.Point(130, 295); $comboPos.Size = New-Object System.Drawing.Size(295, 26)
$leftPanel.Controls.Add($comboPos)

# Z-Order Layer Combo
$lblLayer = New-Object System.Windows.Forms.Label
$lblLayer.Text = "Z-Order Layer:"; $lblLayer.Location = New-Object System.Drawing.Point(12, 335); $lblLayer.AutoSize = $true; $leftPanel.Controls.Add($lblLayer)
$comboLayer = New-Object System.Windows.Forms.ComboBox
$comboLayer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$null = $comboLayer.Items.AddRange(@("Desktop Background (Behind Apps)", "Always On Top"))
$comboLayer.SelectedIndex = 0
$comboLayer.Location = New-Object System.Drawing.Point(130, 332); $comboLayer.Size = New-Object System.Drawing.Size(295, 26)
$leftPanel.Controls.Add($comboLayer)

# Section 4: System Fields Toggle
Add-SectionLabel $leftPanel "4. Visible Diagnostic Fields" 375

$chkHost = New-Object System.Windows.Forms.CheckBox
$chkHost.Text = "Host Name"; $chkHost.Checked = $true; $chkHost.Location = New-Object System.Drawing.Point(16, 405); $chkHost.AutoSize = $true; $leftPanel.Controls.Add($chkHost)

$chkUser = New-Object System.Windows.Forms.CheckBox
$chkUser.Text = "User Name"; $chkUser.Checked = $true; $chkUser.Location = New-Object System.Drawing.Point(120, 405); $chkUser.AutoSize = $true; $leftPanel.Controls.Add($chkUser)

$chkIP = New-Object System.Windows.Forms.CheckBox
$chkIP.Text = "IPv4 Address"; $chkIP.Checked = $true; $chkIP.Location = New-Object System.Drawing.Point(225, 405); $chkIP.AutoSize = $true; $leftPanel.Controls.Add($chkIP)

$chkSerial = New-Object System.Windows.Forms.CheckBox
$chkSerial.Text = "Serial Number"; $chkSerial.Checked = $true; $chkSerial.Location = New-Object System.Drawing.Point(330, 405); $chkSerial.AutoSize = $true; $leftPanel.Controls.Add($chkSerial)

# --- RIGHT PANEL: LIVE PREVIEW & ACTIONS ---
Add-SectionLabel $rightPanel "Live Interactive Preview" 12

# Preview Box Container
$previewBox = New-Object System.Windows.Forms.Panel
$previewBox.Location  = New-Object System.Drawing.Point(16, 42)
$previewBox.Size      = New-Object System.Drawing.Size(378, 230)
$previewBox.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$previewBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$rightPanel.Controls.Add($previewBox)

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

# Function to Update Preview Card Text & Colors
function Update-Preview {
    $Config.Title       = $txtTitle.Text
    $Config.Phone       = $txtPhone.Text
    $Config.Email       = $txtEmail.Text
    $Config.Opacity     = ($trackOpacity.Value / 100.0)
    $Config.Position    = $comboPos.SelectedItem
    $Config.AlwaysOnTop = ($comboLayer.SelectedIndex -eq 1)
    $Config.ShowHost    = $chkHost.Checked
    $Config.ShowUser    = $chkUser.Checked
    $Config.ShowIP      = $chkIP.Checked
    $Config.ShowSerial  = $chkSerial.Checked

    $lblOpacity.Text = "Card Opacity ($($trackOpacity.Value)%):"

    # Color Updates
    $previewCard.BackColor   = $Config.BgColor
    $previewAccent.BackColor = $Config.AccentColor
    $previewLabel.ForeColor  = $Config.TextColor

    # Build Preview String
    $lines = @()
    $lines += $Config.Title.ToUpper()
    $lines += "-----------------------------------------------"
    $lines += "Phone:  $($Config.Phone)"
    $lines += "Email:  $($Config.Email)"
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

# Attach Live Update Events
$txtTitle.Add_TextChanged({ Update-Preview })
$txtPhone.Add_TextChanged({ Update-Preview })
$txtEmail.Add_TextChanged({ Update-Preview })
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
function Save-OverlayDefaults {
    $srcFile = Join-Path -Path $Script:RootPath -ChildPath "src\Start-ITOverlay.ps1"
    if (Test-Path -Path $srcFile) {
        $content = Get-Content -Path $srcFile -Raw
        
        $accentHex = "#{0:X2}{1:X2}{2:X2}" -f $Config.AccentColor.R, $Config.AccentColor.G, $Config.AccentColor.B
        $bgHex     = "#{0:X2}{1:X2}{2:X2}" -f $Config.BgColor.R, $Config.BgColor.G, $Config.BgColor.B
        $textHex   = "#{0:X2}{1:X2}{2:X2}" -f $Config.TextColor.R, $Config.TextColor.G, $Config.TextColor.B

        $content = $content -replace '\[string\]\$Title\s*=\s*"[^"]*"', ('[string]$Title = "{0}"' -f $Config.Title)
        $content = $content -replace '\[string\]\$SupportPhone\s*=\s*"[^"]*"', ('[string]$SupportPhone = "{0}"' -f $Config.Phone)
        $content = $content -replace '\[string\]\$SupportEmail\s*=\s*"[^"]*"', ('[string]$SupportEmail = "{0}"' -f $Config.Email)
        $content = $content -replace '\[string\]\$Position\s*=\s*"[^"]*"', ('[string]$Position = "{0}"' -f $Config.Position)
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

        Set-Content -Path $srcFile -Value $content -Encoding UTF8
    }
}

# --- ACTION BUTTONS ---

# 1. Test Live Overlay Button
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text      = "Test Live Overlay on Desktop"
$btnTest.Location  = New-Object System.Drawing.Point(16, 300)
$btnTest.Size      = New-Object System.Drawing.Size(378, 46)
$btnTest.BackColor = [System.Drawing.Color]::FromArgb(14, 165, 233)
$btnTest.ForeColor = [System.Drawing.Color]::White
$btnTest.Font      = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$btnTest.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rightPanel.Controls.Add($btnTest)

$btnTest.Add_Click({
    try {
        Update-Preview
        $ScriptPath = Join-Path -Path $Script:RootPath -ChildPath "src\Start-ITOverlay.ps1"
        $accentHex = "#{0:X2}{1:X2}{2:X2}" -f $Config.AccentColor.R, $Config.AccentColor.G, $Config.AccentColor.B
        $bgHex     = "#{0:X2}{1:X2}{2:X2}" -f $Config.BgColor.R, $Config.BgColor.G, $Config.BgColor.B
        $textHex   = "#{0:X2}{1:X2}{2:X2}" -f $Config.TextColor.R, $Config.TextColor.G, $Config.TextColor.B

        $hostVal   = if ($Config.ShowHost) { 1 } else { 0 }
        $userVal   = if ($Config.ShowUser) { 1 } else { 0 }
        $ipVal     = if ($Config.ShowIP) { 1 } else { 0 }
        $serialVal = if ($Config.ShowSerial) { 1 } else { 0 }

        $argList = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ScriptPath,
            "-Title", $Config.Title,
            "-SupportPhone", $Config.Phone,
            "-SupportEmail", $Config.Email,
            "-Position", $Config.Position,
            "-AccentColorHex", $accentHex,
            "-BgColorHex", $bgHex,
            "-TextColorHex", $textHex,
            "-ShowHost", $hostVal,
            "-ShowUser", $userVal,
            "-ShowIP", $ipVal,
            "-ShowSerial", $serialVal
        )
        # Always float topmost during live studio testing so test overlay is visible on screen
        $argList += "-AlwaysOnTop"

        Start-Process -FilePath "powershell.exe" -ArgumentList $argList
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error testing live overlay:`n`n$_", "Test Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# 2. Export & Build Intune Package Button
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text      = "Export and Build Intune Package (.intunewim)"
$btnExport.Location  = New-Object System.Drawing.Point(16, 362)
$btnExport.Size      = New-Object System.Drawing.Size(378, 52)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.Font      = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$rightPanel.Controls.Add($btnExport)

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(16, 430)
$lblStatus.Size      = New-Object System.Drawing.Size(378, 80)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$rightPanel.Controls.Add($lblStatus)

$btnExport.Add_Click({
    try {
        Update-Preview
        $lblStatus.Text = "Saving custom defaults and building Intune package..."
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(56, 189, 248)
        $studio.Refresh()

        # 1. Update src/Start-ITOverlay.ps1 defaults
        Save-OverlayDefaults

        # 2. Run Build-IntunePackage.ps1
        $builderScript = Join-Path -Path $Script:RootPath -ChildPath "Build-IntunePackage.ps1"
        $p = Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$builderScript`"" -PassThru -NoNewWindow -Wait

        if ($p.ExitCode -eq 0) {
            $outFile = Join-Path -Path $Script:RootPath -ChildPath "output\Install-ITOverlay.intunewin"
            if (Test-Path -Path $outFile) {
                $lblStatus.Text = "SUCCESS! Package generated:`n" + $outFile
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
                [System.Windows.Forms.MessageBox]::Show("Custom design saved & Intune Win32 Package (.intunewim) compiled successfully!`n`nLocation:`n" + $outFile, "Build Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            $lblStatus.Text = "Error compiling .intunewim package. Exit code: " + $p.ExitCode
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
        }
    } catch {
        $lblStatus.Text = "Error during export: $_"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
        [System.Windows.Forms.MessageBox]::Show("Error during export:`n`n$_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Initial Render
Update-Preview

# Show Application
[System.Windows.Forms.Application]::Run($studio)
