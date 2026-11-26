<#
.SYNOPSIS
    Identifies and removes orphaned Active Directory and Azure AD / Entra ID user profiles on a PC.
.DESCRIPTION
    This script scans local user profiles, filters out system, local, and active profiles, and deletes orphaned AD & Azure AD profiles.
    Saves a persistent log to the Logs directory and runs in WhatIf (dry-run) mode first.
.PARAMETER Force
    Runs the cleanup silently without confirmation prompts.
.PARAMETER ExcludeUsers
    An array of usernames (domain\user) to exclude from the cleanup.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$Force,
    [switch]$DryRun,
    [switch]$ConsoleMode,
    [string[]]$ExcludeUsers
)

# Define Log Path in UK Date Format (DDMMYY)
$UKDate = (Get-Date).ToString("ddMMyy")
$LogDirectory = "$env:SystemDrive\Logs\RemoveOldUserProfiles"
$LogPath = Join-Path -Path $LogDirectory -ChildPath "RemoveOldUserProfiles_$UKDate.log"

# Ensure the log folder exists
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [ConsoleColor]$ForegroundColor = "White"
    )
    $Timestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $LogLine = "[$Timestamp] [$Level] $Message"
    try {
        Add-Content -Path $LogPath -Value $LogLine -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to append to log file: $($_.Exception.Message)"
    }
    
    $Color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        default   { $ForegroundColor }
    }
    Write-Host $LogLine -ForegroundColor $Color
}

# 1. Gain 'System' level access permissions
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Please run as Administrator!" "ERROR"
    return
}

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

# Helper to query profiles
function Get-UserProfileList {
    $fso = New-Object -ComObject Scripting.FileSystemObject
    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    
    $regItems = @()
    try {
        $regItems = Get-ChildItem -Path $regPath -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to read ProfileList registry path: $($_.Exception.Message)"
    }
    
    $regItems |
        Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)-' } |
        ForEach-Object {
            $sid = $_.PSChildName
            $profilePath = $_.GetValue("ProfileImagePath")
            if (-not [string]::IsNullOrWhiteSpace($profilePath)) {
                $folderName = Split-Path $profilePath -Leaf
                $username = if ($sid -match '^S-1-12-1-') { "AzureAD\$folderName" } else { "Unknown SID ($folderName)" }
            } else {
                $username = if ($sid -match '^S-1-12-1-') { "AzureAD\$sid [Path Missing]" } else { "Unknown SID ($sid) [Path Missing]" }
            }
            
            try {
                $account = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount])
                $username = $account.Value
            } catch {}

            $size = 0
            if ($profilePath -and (Test-Path -Path $profilePath -PathType Container)) {
                try {
                    $fsoSize = $fso.GetFolder($profilePath).Size
                    if ($null -ne $fsoSize) { $size = $fsoSize }
                } catch {}

                # Fallback if COM object silently failed or returned null
                if ($size -eq 0) {
                    try { 
                        $fallbackSize = (Get-ChildItem -LiteralPath $profilePath -Recurse -File -Force -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum
                        if ($null -ne $fallbackSize) { $size = $fallbackSize }
                    } catch {}
                }
            }

            $isActive = Test-Path "Registry::HKEY_USERS\$sid"
            try {
                $consoleUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
                if ($consoleUser -and $username -eq $consoleUser) {
                    $isActive = $true
                }
            } catch {}

            [PSCustomObject]@{
                SID         = $sid
                ProfilePath = $profilePath
                UserName    = $username
                SizeGB      = [math]::Round($size / 1GB, 3)
                IsActive    = $isActive
            }
        }
}

# 1. Gather profiles
$currentLoggedInUserSID = ""
try {
    $currentLoggedInUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
} catch {}

$wellKnownSids = @()
if ($ExcludeUsers) {
    foreach ($user in $ExcludeUsers) {
        try {
            $wellKnownSids += (New-Object System.Security.Principal.NTAccount($user)).Translate([System.Security.Principal.UIDentifier]).Value
        } catch {}
    }
}

Write-Log "Scanning local user profiles..." "INFO" "Cyan"
$profiles = Get-UserProfileList
$profilesToRemove = @()

foreach ($profile in $profiles) {
    # Skip current logged in user
    if ($profile.SID -eq $currentLoggedInUserSID) { continue }
    # Skip explicitly excluded users
    if ($profile.SID -in $wellKnownSids) { continue }
    
    # Check if old domain account (domain\username where domain != computername and BUILTIN)
    if ($profile.UserName -like "*\*") {
        $parts = $profile.UserName.Split('\')
        $domain = $parts[0]
        if ($domain -ne $env:COMPUTERNAME -and $domain -ne "BUILTIN") {
            $profilesToRemove += $profile
        }
    } else {
        # Unresolved SIDs are likely deleted/orphaned domain accounts
        $profilesToRemove += $profile
    }
}

if ($profilesToRemove.Count -eq 0) {
    Write-Log "No orphaned AD / Azure AD user profiles found." "SUCCESS"
    return
}

Write-Log "Found $($profilesToRemove.Count) orphaned AD / Azure AD profiles." "WARNING"
$profilesToRemove | Format-Table UserName, ProfilePath, SizeGB -AutoSize | Out-String | ForEach-Object {
    if (-not [string]::IsNullOrWhiteSpace($_)) {
        Write-Host $_ -ForegroundColor Yellow
        try {
            Add-Content -Path $LogPath -Value $_ -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to append table output to log file: $($_.Exception.Message)"
        }
    }
}

function Invoke-ProfileCleanup {
    foreach ($profile in $profilesToRemove) {
        try {
            Write-Log "Removing profile for $($profile.UserName)..." "INFO" "Cyan"
            if ($WhatIfPreference) {
                Write-Host "What if: Performing cleanup on profile $($profile.ProfilePath) (SID: $($profile.SID))" -ForegroundColor Gray
            } else {
                $deleted = $false
                $cimProfile = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($profile.SID)'"
                if ($cimProfile) {
                    Remove-CimInstance -InputObject $cimProfile -ErrorAction Stop
                    Write-Log "Successfully removed profile via CIM." "SUCCESS"
                    $deleted = $true
                } else {
                    Write-Log "Profile not found via CIM. Attempting forced manual cleanup..." "WARNING"
                }

                # Force Cleanup Fallback for Ghost SIDs
                if (-not $deleted) {
                    # 1. Delete physical folder if it exists
                    if ($profile.ProfilePath -and (Test-Path -Path $profile.ProfilePath)) {
                        try {
                            Remove-Item -Path $profile.ProfilePath -Force -Recurse -ErrorAction Stop
                            Write-Log "Force deleted stranded profile folder." "SUCCESS"
                        } catch {
                            Write-Log "Partial or failed deletion of profile folder: $($_.Exception.Message)" "WARNING"
                        }
                    }
                    
                    # 2. Delete the dead registry key
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($profile.SID)"
                    if (Test-Path -Path $regPath) {
                        try {
                            Remove-Item -Path $regPath -Force -Recurse -ErrorAction Stop
                            Write-Log "Successfully purged dead registry key." "SUCCESS"
                        } catch {
                            Write-Log "Failed to purge dead registry key: $($_.Exception.Message)" "WARNING"
                        }
                    }
                }
            }
        } catch {
            Write-Log "Failed to remove profile: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Initialize-StandardWinForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form,
        [string]$ScriptRoot = $PSScriptRoot,
        [bool]$IsDark = $true
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

    $Script:UiFontCollection = New-Object System.Drawing.Text.PrivateFontCollection
    $figtreeFamily = $null
    if ($resolvedFontDir) {
        foreach ($fontFile in $fontFiles) {
            $fontPath = Join-Path -Path $resolvedFontDir -ChildPath $fontFile
            if (Test-Path -LiteralPath $fontPath) {
                try {
                    $Script:UiFontCollection.AddFontFile($fontPath)
                } catch {}
            }
        }
        foreach ($family in $Script:UiFontCollection.Families) {
            if ($family.Name -eq 'Figtree') {
                $figtreeFamily = $family
                break
            }
        }
    }

    $makeFont = {
        param([float]$Size, [System.Drawing.FontStyle]$Style)
        if ($figtreeFamily) {
            try {
                return New-Object System.Drawing.Font($figtreeFamily, $Size, $Style)
            } catch {}
        }
        return New-Object System.Drawing.Font('Segoe UI', $Size, $Style)
    }

    $workArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Form.Size = New-Object System.Drawing.Size($targetWidth, $targetHeight)
    $Form.MinimumSize = New-Object System.Drawing.Size([Math]::Min(760, $targetWidth), [Math]::Min(520, $targetHeight))
    $Form.Location = New-Object System.Drawing.Point(
        ($workArea.Left + [Math]::Round(($workArea.Width - $targetWidth) / 2)),
        ($workArea.Top  + [Math]::Round(($workArea.Height - $targetHeight) / 2))
    )
    if ($IsDark) {
        $Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#091122')
        $Form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#F8FAFC')
    } else {
        $Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F8F9FF')
        $Form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
    }

    $bodyFont = & $makeFont 10 ([System.Drawing.FontStyle]::Regular)
    $Form.Font = $bodyFont

    return @{
        Title      = (& $makeFont 15 ([System.Drawing.FontStyle]::Bold))
        Subtitle   = (& $makeFont 9.5 ([System.Drawing.FontStyle]::Regular))
        Section    = (& $makeFont 8.5 ([System.Drawing.FontStyle]::Bold))
        Body       = $bodyFont
        Button     = (& $makeFont 9.5 ([System.Drawing.FontStyle]::Regular))
        ButtonBold = (& $makeFont 9.5 ([System.Drawing.FontStyle]::Bold))
        Mono       = (New-Object System.Drawing.Font('Consolas', 8.5, ([System.Drawing.FontStyle]::Regular)))
    }
}

# Execution Flow
if (-not [Environment]::UserInteractive) {
    Write-Log "Non-interactive session detected. Forcing Console Mode." "WARNING" "Yellow"
    $ConsoleMode = $true
}

if ($Force) {
    $WhatIfPreference = $false
    Invoke-ProfileCleanup
} elseif ($DryRun) {
    Write-Log "=== STARTING WHATIF DRY-RUN (Strict Mode) ===" "WARNING"
    $WhatIfPreference = $true
    Invoke-ProfileCleanup
    Write-Log "Strict DryRun completed. No changes were made." "INFO" "Gray"
} elseif ($ConsoleMode) {
    Write-Log "Launching interactive console selection menu..." "INFO" "Cyan"
    Write-Host "`nAll profiles below will be permanently deleted. Select the ones you want to delete:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $profilesToRemove.Count; $i++) {
        $profile = $profilesToRemove[$i]
        $activeTag = if ($profile.IsActive) { " [ACTIVE USER]" } else { "" }
        Write-Host "[$i] $($profile.UserName)$activeTag | Size: $($profile.SizeGB) GB"
    }
    
    Write-Host ""
    $selection = Read-Host "Enter numbers to DELETE separated by commas (e.g. 0,2,3), or 'all' to delete everything, or 'none' to cancel"
    
    if ($selection -eq 'none' -or [string]::IsNullOrWhiteSpace($selection)) {
        Write-Log "Cancelled by user. No changes were made." "INFO" "Gray"
    } elseif ($selection -eq 'all') {
        Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
        $WhatIfPreference = $false
        Invoke-ProfileCleanup
    } else {
        $approvedProfiles = @()
        $indexes = $selection -split ',' | ForEach-Object { $_.Trim() }
        foreach ($index in $indexes) {
            if ($index -match '^\d+$') {
                $idx = [int]$index
                if ($idx -ge 0 -and $idx -lt $profilesToRemove.Count) {
                    $approvedProfiles += $profilesToRemove[$idx]
                }
            }
        }
        
        $profilesToRemove = $approvedProfiles
        if ($profilesToRemove.Count -eq 0) {
            Write-Log "No valid profiles selected. No changes were made." "INFO" "Gray"
        } else {
            Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
            $WhatIfPreference = $false
            Invoke-ProfileCleanup
        }
    }
} else {
    Write-Log "Launching interactive profile selection menu..." "INFO" "Cyan"
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Script:IsDarkTheme = Get-SystemIsDarkTheme
    if ($Script:IsDarkTheme) {
        $colorHeaderFooter = [System.Drawing.ColorTranslator]::FromHtml('#070D19')
        $colorCanvas       = [System.Drawing.ColorTranslator]::FromHtml('#091122')
        $colorListBg       = [System.Drawing.ColorTranslator]::FromHtml('#0B1528')
        $colorBorder       = [System.Drawing.ColorTranslator]::FromHtml('#26354D')
        $colorTextPrimary  = [System.Drawing.ColorTranslator]::FromHtml('#F8FAFC')
        $colorTextMuted    = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
        $colorAccent       = [System.Drawing.ColorTranslator]::FromHtml('#2DD4BF')
        $colorSecondaryBtn = [System.Drawing.ColorTranslator]::FromHtml('#13223A')
        $colorSecondaryFg  = [System.Drawing.ColorTranslator]::FromHtml('#E2E8F0')
        $colorSecondaryHov = [System.Drawing.ColorTranslator]::FromHtml('#1E3252')
        $colorDangerBg     = [System.Drawing.ColorTranslator]::FromHtml('#7F1D1D')
        $colorDangerHov    = [System.Drawing.ColorTranslator]::FromHtml('#991B1B')
        $colorDangerBorder = [System.Drawing.ColorTranslator]::FromHtml('#EF4444')
        $colorDangerText   = [System.Drawing.ColorTranslator]::FromHtml('#FEE2E2')
    } else {
        $colorHeaderFooter = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
        $colorCanvas       = [System.Drawing.ColorTranslator]::FromHtml('#F8F9FF')
        $colorListBg       = [System.Drawing.ColorTranslator]::FromHtml('#F8FAFC')
        $colorBorder       = [System.Drawing.ColorTranslator]::FromHtml('#CBD5E1')
        $colorTextPrimary  = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
        $colorTextMuted    = [System.Drawing.ColorTranslator]::FromHtml('#475569')
        $colorAccent       = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
        $colorSecondaryBtn = [System.Drawing.ColorTranslator]::FromHtml('#F1F5F9')
        $colorSecondaryFg  = [System.Drawing.ColorTranslator]::FromHtml('#13223A')
        $colorSecondaryHov = [System.Drawing.ColorTranslator]::FromHtml('#E2E8F0')
        $colorDangerBg     = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
        $colorDangerHov    = [System.Drawing.ColorTranslator]::FromHtml('#B91C1C')
        $colorDangerBorder = [System.Drawing.ColorTranslator]::FromHtml('#B91C1C')
        $colorDangerText   = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Remove Old User Profiles'
    $form.Topmost = $true
    $uiFonts = Initialize-StandardWinForm -Form $form -ScriptRoot $PSScriptRoot -IsDark $Script:IsDarkTheme
    $form.Add_HandleCreated({ try { if ('Win11ChromeHelper' -as [type]) { [Win11ChromeHelper]::ApplyWindowChrome($form.Handle, $Script:IsDarkTheme) } } catch {} })

    # Header Banner
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $headerPanel.Height = 76
    $headerPanel.BackColor = $colorHeaderFooter
    $headerPanel.Padding = New-Object System.Windows.Forms.Padding(22, 14, 22, 12)

    $headerBorder = New-Object System.Windows.Forms.Panel
    $headerBorder.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $headerBorder.Height = 1
    $headerBorder.BackColor = $colorBorder

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $titleLabel.Height = 28
    $titleLabel.Text = 'Remove Old User Profiles'
    $titleLabel.Font = $uiFonts.Title
    $titleLabel.ForeColor = $colorTextPrimary
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $subtitleLabel.Text = 'All profiles below will be permanently deleted. Untick any you want to keep.'
    $subtitleLabel.Font = $uiFonts.Subtitle
    $subtitleLabel.ForeColor = $colorTextMuted
    $subtitleLabel.BackColor = [System.Drawing.Color]::Transparent

    $headerPanel.Controls.Add($subtitleLabel)
    $headerPanel.Controls.Add($titleLabel)
    $headerPanel.Controls.Add($headerBorder)
    $subtitleLabel.BringToFront()

    # Footer Action Bar
    $footerPanel = New-Object System.Windows.Forms.Panel
    $footerPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footerPanel.Height = 64
    $footerPanel.BackColor = $colorHeaderFooter
    $footerPanel.Padding = New-Object System.Windows.Forms.Padding(22, 12, 22, 12)

    $footerBorder = New-Object System.Windows.Forms.Panel
    $footerBorder.Dock = [System.Windows.Forms.DockStyle]::Top
    $footerBorder.Height = 1
    $footerBorder.BackColor = $colorBorder

    $footerLogLabel = New-Object System.Windows.Forms.Label
    $footerLogLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $footerLogLabel.Text = "Log: $LogPath"
    $footerLogLabel.Font = $uiFonts.Mono
    $footerLogLabel.ForeColor = $colorTextMuted
    $footerLogLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $footerLogLabel.AutoEllipsis = $true

    $buttonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonFlow.Dock = [System.Windows.Forms.DockStyle]::Right
    $buttonFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonFlow.WrapContents = $false
    $buttonFlow.AutoSize = $true
    $buttonFlow.BackColor = [System.Drawing.Color]::Transparent
    $buttonFlow.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Size = New-Object System.Drawing.Size(96, 34)
    $selectAllButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $selectAllButton.Text = 'Select All'
    $selectAllButton.Font = $uiFonts.Button
    $selectAllButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $selectAllButton.BackColor = $colorSecondaryBtn
    $selectAllButton.ForeColor = $colorSecondaryFg
    $selectAllButton.FlatAppearance.BorderColor = $colorBorder
    $selectAllButton.FlatAppearance.BorderSize = 1
    $selectAllButton.FlatAppearance.MouseOverBackColor = $colorSecondaryHov
    $selectAllButton.Cursor = [System.Windows.Forms.Cursors]::Hand

    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Size = New-Object System.Drawing.Size(84, 34)
    $clearButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $clearButton.Text = 'Clear'
    $clearButton.Font = $uiFonts.Button
    $clearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $clearButton.BackColor = $colorSecondaryBtn
    $clearButton.ForeColor = $colorSecondaryFg
    $clearButton.FlatAppearance.BorderColor = $colorBorder
    $clearButton.FlatAppearance.BorderSize = 1
    $clearButton.FlatAppearance.MouseOverBackColor = $colorSecondaryHov
    $clearButton.Cursor = [System.Windows.Forms.Cursors]::Hand

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(88, 34)
    $cancelButton.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $cancelButton.Text = 'Cancel'
    $cancelButton.Font = $uiFonts.Button
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.BackColor = $colorSecondaryBtn
    $cancelButton.ForeColor = $colorSecondaryFg
    $cancelButton.FlatAppearance.BorderColor = $colorBorder
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.FlatAppearance.MouseOverBackColor = $colorSecondaryHov
    $cancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(134, 34)
    $okButton.Margin = New-Object System.Windows.Forms.Padding(0)
    $okButton.Text = 'Delete Selected'
    $okButton.Font = $uiFonts.ButtonBold
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $okButton.BackColor = $colorDangerBg
    $okButton.ForeColor = $colorDangerText
    $okButton.FlatAppearance.BorderColor = $colorDangerBorder
    $okButton.FlatAppearance.BorderSize = 1
    $okButton.FlatAppearance.MouseOverBackColor = $colorDangerHov
    $okButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $buttonFlow.Controls.Add($selectAllButton)
    $buttonFlow.Controls.Add($clearButton)
    $buttonFlow.Controls.Add($cancelButton)
    $buttonFlow.Controls.Add($okButton)

    $footerPanel.Controls.Add($footerLogLabel)
    $footerPanel.Controls.Add($buttonFlow)
    $footerPanel.Controls.Add($footerBorder)
    $footerLogLabel.BringToFront()

    # Workspace Body
    $bodyPanel = New-Object System.Windows.Forms.Panel
    $bodyPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $bodyPanel.BackColor = $colorCanvas
    $bodyPanel.Padding = New-Object System.Windows.Forms.Padding(22, 16, 22, 16)

    $sectionLabel = New-Object System.Windows.Forms.Label
    $sectionLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $sectionLabel.Height = 24
    $sectionLabel.Text = "ORPHANED PROFILES DETECTED ($($profilesToRemove.Count))"
    $sectionLabel.Font = $uiFonts.Section
    $sectionLabel.ForeColor = $colorAccent

    $listBorderPanel = New-Object System.Windows.Forms.Panel
    $listBorderPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $listBorderPanel.BackColor = $colorBorder
    $listBorderPanel.Padding = New-Object System.Windows.Forms.Padding(1)

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $checkedListBox.BackColor = $colorListBg
    $checkedListBox.ForeColor = $colorTextPrimary
    $checkedListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.IntegralHeight = $false
    $checkedListBox.Font = $uiFonts.Body

    foreach ($profile in $profilesToRemove) {
        $activeTag = if ($profile.IsActive) { " [ACTIVE USER]" } else { "" }
        $item = "$($profile.UserName)$activeTag | Size: $($profile.SizeGB) GB"
        $isChecked = -not $profile.IsActive
        $checkedListBox.Items.Add($item, $isChecked) | Out-Null
    }

    $listBorderPanel.Controls.Add($checkedListBox)
    $bodyPanel.Controls.Add($listBorderPanel)
    $bodyPanel.Controls.Add($sectionLabel)
    $listBorderPanel.BringToFront()

    $form.Controls.Add($bodyPanel)
    $form.Controls.Add($footerPanel)
    $form.Controls.Add($headerPanel)
    $bodyPanel.BringToFront()

    $selectAllButton.Add_Click({
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            $checkedListBox.SetItemChecked($i, $true)
        }
    })

    $clearButton.Add_Click({
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            $checkedListBox.SetItemChecked($i, $false)
        }
    })

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $approvedProfiles = @()
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            if ($checkedListBox.GetItemChecked($i)) {
                $approvedProfiles += $profilesToRemove[$i]
            }
        }

        $profilesToRemove = $approvedProfiles
        if ($profilesToRemove.Count -eq 0) {
            Write-Log "All profiles were unchecked. No changes were made." "INFO" "Gray"
        } else {
            Write-Log "=== RUNNING REAL CLEANUP ===" "SUCCESS"
            $WhatIfPreference = $false
            Invoke-ProfileCleanup
        }
    } else {
        Write-Log "Cancelled by user. No changes were made." "INFO" "Gray"
    }
}
