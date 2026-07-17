<#
.SYNOPSIS
    GUI utility for building, staging, and packaging corporate Desktop Background and Lock Screen deployments for Intune.
.DESCRIPTION
    This script launches a Windows Forms user interface to configure corporate branding assets. It supports local image files or remote URLs, and generates staging folders, installation scripts (Install.ps1, ApplyBackground.ps1, DownloadBackgrounds.ps1), custom detection scripts, and packs everything into a ready-to-upload .intunewin bundle.
.PARAMETER DebugMode
    Launches the packaging process with verbose execution logs.
.NOTES
    Author     : JP
    Created    : 2026-07-17
    Version    : 1.0
    Log Path   : C:\Logs\DesktopBackground\
#>
param(
    [switch]$DebugMode = $false
)

# Add necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- UI Configuration ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Desktop Background Packager"
$form.Size = New-Object System.Drawing.Size(600, 450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White

# Fonts
$titleFont = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$labelFont = New-Object System.Drawing.Font("Segoe UI", 10)
$textFont = New-Object System.Drawing.Font("Segoe UI", 9)

# Title
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Intune Background Packager"
$lblTitle.Font = $titleFont
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.AutoSize = $true
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$form.Controls.Add($lblTitle)

# Group Box: Source Type
$grpSource = New-Object System.Windows.Forms.GroupBox
$grpSource.Text = "Desktop Background"
$grpSource.Font = $labelFont
$grpSource.Location = New-Object System.Drawing.Point(20, 70)
$grpSource.Size = New-Object System.Drawing.Size(540, 175)
$form.Controls.Add($grpSource)

# Enable Desktop Background Checkbox
$chkSetDesktop = New-Object System.Windows.Forms.CheckBox
$chkSetDesktop.Text = "Enable Desktop Background Management"
$chkSetDesktop.Checked = $true
$chkSetDesktop.Location = New-Object System.Drawing.Point(20, 25)
$chkSetDesktop.AutoSize = $true
$grpSource.Controls.Add($chkSetDesktop)

# Radio Buttons
$radLocal = New-Object System.Windows.Forms.RadioButton
$radLocal.Text = "Local Image File"
$radLocal.Location = New-Object System.Drawing.Point(20, 55)
$radLocal.Checked = $true
$radLocal.AutoSize = $true
$grpSource.Controls.Add($radLocal)

$radUrl = New-Object System.Windows.Forms.RadioButton
$radUrl.Text = "Image URL (Dynamic Update)"
$radUrl.Location = New-Object System.Drawing.Point(20, 95)
$radUrl.AutoSize = $true
$grpSource.Controls.Add($radUrl)

# Input for Local File
$txtLocalFile = New-Object System.Windows.Forms.TextBox
$txtLocalFile.Location = New-Object System.Drawing.Point(230, 53)
$txtLocalFile.Size = New-Object System.Drawing.Size(200, 25)
$txtLocalFile.Font = $textFont
$grpSource.Controls.Add($txtLocalFile)

$btnBrowseLocal = New-Object System.Windows.Forms.Button
$btnBrowseLocal.Text = "Browse..."
$btnBrowseLocal.Location = New-Object System.Drawing.Point(440, 52)
$btnBrowseLocal.Size = New-Object System.Drawing.Size(80, 27)
$btnBrowseLocal.Font = $textFont
$grpSource.Controls.Add($btnBrowseLocal)

# Input for URL
$txtUrl = New-Object System.Windows.Forms.TextBox
$txtUrl.Location = New-Object System.Drawing.Point(230, 95)
$txtUrl.Size = New-Object System.Drawing.Size(290, 25)
$txtUrl.Font = $textFont
$txtUrl.Enabled = $false
$grpSource.Controls.Add($txtUrl)

# Prevent Changing Checkbox
$chkLockBackground = New-Object System.Windows.Forms.CheckBox
$chkLockBackground.Text = "Prevent users from changing the Desktop Background"
$chkLockBackground.Location = New-Object System.Drawing.Point(20, 135)
$chkLockBackground.AutoSize = $true
$grpSource.Controls.Add($chkLockBackground)

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
    }
})

# Group Box: Lock Screen Source
$grpLockScreen = New-Object System.Windows.Forms.GroupBox
$grpLockScreen.Text = "Lock Screen"
$grpLockScreen.Font = $labelFont
$grpLockScreen.Location = New-Object System.Drawing.Point(20, 260)
$grpLockScreen.Size = New-Object System.Drawing.Size(540, 185)
$form.Controls.Add($grpLockScreen)

# Enable Lock Screen Checkbox
$chkSetLockScreen = New-Object System.Windows.Forms.CheckBox
$chkSetLockScreen.Text = "Enable Lock Screen Management"
$chkSetLockScreen.Checked = $true
$chkSetLockScreen.Location = New-Object System.Drawing.Point(20, 25)
$chkSetLockScreen.AutoSize = $true
$grpLockScreen.Controls.Add($chkSetLockScreen)

$radLockSame = New-Object System.Windows.Forms.RadioButton
$radLockSame.Text = "Same as Desktop Background"
$radLockSame.Location = New-Object System.Drawing.Point(20, 55)
$radLockSame.Checked = $true
$radLockSame.AutoSize = $true
$grpLockScreen.Controls.Add($radLockSame)

$radLockLocal = New-Object System.Windows.Forms.RadioButton
$radLockLocal.Text = "Local Image File"
$radLockLocal.Location = New-Object System.Drawing.Point(20, 85)
$radLockLocal.AutoSize = $true
$grpLockScreen.Controls.Add($radLockLocal)

$radLockUrl = New-Object System.Windows.Forms.RadioButton
$radLockUrl.Text = "Image URL (Dynamic Update)"
$radLockUrl.Location = New-Object System.Drawing.Point(20, 115)
$radLockUrl.AutoSize = $true
$grpLockScreen.Controls.Add($radLockUrl)

$txtLockLocalFile = New-Object System.Windows.Forms.TextBox
$txtLockLocalFile.Location = New-Object System.Drawing.Point(230, 83)
$txtLockLocalFile.Size = New-Object System.Drawing.Size(200, 25)
$txtLockLocalFile.Font = $textFont
$txtLockLocalFile.Enabled = $false
$grpLockScreen.Controls.Add($txtLockLocalFile)

$btnBrowseLockLocal = New-Object System.Windows.Forms.Button
$btnBrowseLockLocal.Text = "Browse..."
$btnBrowseLockLocal.Location = New-Object System.Drawing.Point(440, 82)
$btnBrowseLockLocal.Size = New-Object System.Drawing.Size(80, 27)
$btnBrowseLockLocal.Font = $textFont
$btnBrowseLockLocal.Enabled = $false
$grpLockScreen.Controls.Add($btnBrowseLockLocal)

$txtLockUrl = New-Object System.Windows.Forms.TextBox
$txtLockUrl.Location = New-Object System.Drawing.Point(230, 115)
$txtLockUrl.Size = New-Object System.Drawing.Size(290, 25)
$txtLockUrl.Font = $textFont
$txtLockUrl.Enabled = $false
$grpLockScreen.Controls.Add($txtLockUrl)

# Prevent Changing Lock Screen Checkbox
$chkLockLockScreen = New-Object System.Windows.Forms.CheckBox
$chkLockLockScreen.Text = "Prevent users from changing the Lock Screen"
$chkLockLockScreen.Location = New-Object System.Drawing.Point(20, 145)
$chkLockLockScreen.AutoSize = $true
$grpLockScreen.Controls.Add($chkLockLockScreen)

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
    }
})

# Adjust Form Size and Output Group Box Location
$form.Size = New-Object System.Drawing.Size(600, 680)

# Group Box: Output
$grpOutput = New-Object System.Windows.Forms.GroupBox
$grpOutput.Text = "Output Directory"
$grpOutput.Font = $labelFont
$grpOutput.Location = New-Object System.Drawing.Point(20, 460)
$grpOutput.Size = New-Object System.Drawing.Size(540, 70)
$form.Controls.Add($grpOutput)

$txtOutputDir = New-Object System.Windows.Forms.TextBox
$txtOutputDir.Location = New-Object System.Drawing.Point(20, 30)
$txtOutputDir.Size = New-Object System.Drawing.Size(410, 25)
$txtOutputDir.Font = $textFont

$currentScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($currentScriptDir)) { $currentScriptDir = (Get-Location).Path }
$txtOutputDir.Text = Join-Path $currentScriptDir "Output"

$grpOutput.Controls.Add($txtOutputDir)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = "Browse..."
$btnBrowseOutput.Location = New-Object System.Drawing.Point(440, 28)
$btnBrowseOutput.Size = New-Object System.Drawing.Size(80, 27)
$btnBrowseOutput.Font = $textFont
$grpOutput.Controls.Add($btnBrowseOutput)

$btnBrowseOutput.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.SelectedPath = $txtOutputDir.Text
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $folderBrowser.SelectedPath
    }
})

# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 540)
$lblStatus.Size = New-Object System.Drawing.Size(540, 40)
$lblStatus.Font = $textFont
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text = "Ready."
$form.Controls.Add($lblStatus)

# Generate Button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Intune Package"
$btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenerate.Location = New-Object System.Drawing.Point(360, 580)
$btnGenerate.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($btnGenerate)

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
        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = (Get-Location).Path }
        
        $intuneTool = Join-Path $scriptDir "IntuneWinAppUtil.exe"
        if (-not (Test-Path $intuneTool)) {
            $lblStatus.Text = "Downloading IntuneWinAppUtil.exe..."
            $lblStatus.Refresh()
            $toolUrl = "https://raw.githubusercontent.com/microsoft/Microsoft-Win32-Content-Prep-Tool/master/IntuneWinAppUtil.exe"
            Invoke-WebRequest -Uri $toolUrl -OutFile $intuneTool -UseBasicParsing
            if (-not (Test-Path $intuneTool)) {
                throw "Failed to download IntuneWinAppUtil.exe. Please download it manually and place it in the same folder."
            }
        }

        # 3. Setup Staging Directory
        $lblStatus.Text = "Setting up staging directory..."
        $lblStatus.Refresh()
        $stagingDir = Join-Path $scriptDir "Staging_IntuneBackground_Temp"
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
Start-Transcript -Path `"C:\Logs\DesktopBackground\ApplyBackground.txt`" -Append
`$destDesktop = "`$env:ProgramData\CorporateBackground\desktop_$fileName"

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
Start-Transcript -Path `"C:\Logs\DesktopBackground\Install.txt`" -Force
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
            
            if ($setDesktop -and $chkLockBackground.Checked) {
                $installContent += @"

Start-Sleep -Seconds 5
`$ActiveDesktopPath = `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`"
if (!(Test-Path `$ActiveDesktopPath)) { New-Item -Path `$ActiveDesktopPath -Force | Out-Null }
Set-ItemProperty -Path `$ActiveDesktopPath -Name `"NoChangingWallPaper`" -Value 1 -Force
"@
            }

            if ($setLockScreen -and $chkLockLockScreen.Checked) {
                $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"NoChangingLockScreen`" -Value 1 -Force
"@
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
Start-Transcript -Path `"C:\Logs\DesktopBackground\DownloadBackgrounds.txt`" -Append

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

Remove-ItemProperty -Path `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`" -Name `"NoChangingWallPaper`" -ErrorAction SilentlyContinue
"@
            }

            if ($setDesktop) {
                $dlContent += @"

Start-ScheduledTask -TaskName 'ApplyCorporateBackground' -ErrorAction SilentlyContinue
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
Start-Transcript -Path `"C:\Logs\DesktopBackground\ApplyBackground.txt`" -Append
`$destDesktop = "$destDesktopPath"

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
Start-Transcript -Path `"C:\Logs\DesktopBackground\Install.txt`" -Force

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

            if ($setDesktop -and $chkLockBackground.Checked) {
                $installContent += @"

Start-Sleep -Seconds 5
`$ActiveDesktopPath = `"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop`"
if (!(Test-Path `$ActiveDesktopPath)) { New-Item -Path `$ActiveDesktopPath -Force | Out-Null }
Set-ItemProperty -Path `$ActiveDesktopPath -Name `"NoChangingWallPaper`" -Value 1 -Force
"@
            }

            if ($setLockScreen -and $chkLockLockScreen.Checked) {
                $installContent += @"

`$RegPath = `"HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization`"
if (!(Test-Path `$RegPath)) { New-Item -Path `$RegPath -Force | Out-Null }
Set-ItemProperty -Path `$RegPath -Name `"NoChangingLockScreen`" -Value 1 -Force
"@
            }

            $installContent += @"

Stop-Transcript
"@
            Set-Content -Path $installScriptPath -Value $installContent -Force
        }

        # Create Uninstall.ps1
        $uninstallContent = @"
`$ErrorActionPreference = 'Continue'
if (!(Test-Path `"C:\Logs\DesktopBackground`")) { New-Item -ItemType Directory -Path `"C:\Logs\DesktopBackground`" -Force }
Start-Transcript -Path `"C:\Logs\DesktopBackground\Uninstall.txt`" -Append
Remove-Item -Path "`$env:ProgramData\CorporateBackground" -Recurse -Force
Unregister-ScheduledTask -TaskName 'ApplyCorporateBackground' -Confirm:`$false
Unregister-ScheduledTask -TaskName 'CorporateBackgroundDownloader' -Confirm:`$false
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name "LockScreenImage"
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop" -Name "NoChangingWallPaper"
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name "NoChangingLockScreen"
Stop-Transcript
"@
        Set-Content -Path (Join-Path $stagingDir "Uninstall.ps1") -Value $uninstallContent -Force

        # Trim output directory to prevent trailing spaces
        $outDir = $txtOutputDir.Text.Trim()
        
        # Ensure output directory exists
        if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        
        $toolPath = Join-Path $scriptDir "IntuneWinAppUtil.exe"
        $process = Start-Process -FilePath $toolPath -ArgumentList "-c `"$stagingDir`"", "-s `"Install.ps1`"", "-o `"$outDir`"", "-q" -NoNewWindow -PassThru -Wait
        
        if ($process.ExitCode -ne 0) {
            throw "IntuneWinAppUtil failed to package the application. Exit code: $($process.ExitCode)"
        }

        # Rename output .intunewin file to DesktopBackground_DDMMYY.intunewin
        $dateStr = Get-Date -Format "ddMMyy"
        $newWimName = "DesktopBackground_$dateStr.intunewin"
        $wimPath = Join-Path $outDir "Install.intunewin"
        $newWimPath = Join-Path $outDir $newWimName
        
        # Give the filesystem a moment to index the new file
        Start-Sleep -Milliseconds 500
        
        if (Test-Path -LiteralPath $wimPath) {
            if (Test-Path -LiteralPath $newWimPath) { Remove-Item -LiteralPath $newWimPath -Force -ErrorAction SilentlyContinue }
            
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
if ((Test-Path "`$targetPath\ApplyBackground.ps1") -and (Get-ScheduledTask -TaskName 'ApplyCorporateBackground' -ErrorAction SilentlyContinue)) {
    Write-Output "Detected"
    exit 0
} else {
    exit 1
}
"@
            } else {
                # Lock Screen Only (Local)
                $detectionContent = @"
`$targetPath = "`$env:ProgramData\CorporateBackground"
`$regVal = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization" -Name "LockScreenImage" -ErrorAction SilentlyContinue).LockScreenImage
if (`$regVal -and (Test-Path `$regVal)) {
    Write-Output "Detected"
    exit 0
} else {
    exit 1
}
"@
            }
        } else {
            # URL (Desktop or Lock Screen or Both)
            $detectionContent = @"
`$targetPath = "`$env:ProgramData\CorporateBackground"
if ((Test-Path "`$targetPath\DownloadBackgrounds.ps1") -and (Get-ScheduledTask -TaskName 'CorporateBackgroundDownloader' -ErrorAction SilentlyContinue)) {
    Write-Output "Detected"
    exit 0
} else {
    exit 1
}
"@
        }
        $detectionFile = Join-Path $outDir "DetectionScript.ps1"
        Set-Content -Path $detectionFile -Value $detectionContent -Force

        $commandsContent = @"
=== Intune Win32 App Deployment Configuration ===

1. App Information
Name: Corporate Desktop and Lock Screen Background
Description: Sets the desktop background and optionally the lock screen, automatically applying to all users.
Publisher: IT Department

2. Program
Install command: powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install.ps1
Uninstall command: powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall.ps1

Install behavior: System
Device restart behavior: No specific action

3. Requirements
Operating system architecture: x64
Minimum operating system: Windows 10 1607

4. Detection rules
Rules format: Use a custom detection script
Script file: Upload the DetectionScript.ps1 file generated in this folder.
Run script as 32-bit processes on 64-bit clients: No
Enforce script signature check: No
"@
        $commandsFile = Join-Path $outDir "InstallCommands.txt"
        Set-Content -Path $commandsFile -Value $commandsContent -Force

        # 7. Cleanup
        if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }

        $lblStatus.Text = "Package successfully created in Output Directory!"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
        [System.Windows.Forms.MessageBox]::Show("Packaging Complete! Check the output directory for your $newWimName, DetectionScript.ps1, and InstallCommands.txt.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        $lblStatus.Text = "Error: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Show Form
$form.ShowDialog() | Out-Null
$form.Dispose()
