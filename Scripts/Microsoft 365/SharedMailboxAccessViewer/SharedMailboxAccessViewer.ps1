<#
.SYNOPSIS
    Microsoft 365 Shared Mailbox Access Viewer - Interactive WPF GUI Tool (v1.0)

.DESCRIPTION
    Provides a fast, modern PowerShell WPF GUI to audit which Shared Mailboxes (and optionally delegated User Mailboxes)
    a user can access across Exchange Online:
    - Full Access (Get-EXOMailboxPermission)
    - Send As (Get-EXORecipientPermission)
    - Send on Behalf (GrantSendOnBehalfTo)

    Supports two inspection modes:
    1. Single-User Check: Select any user from the left pane and click "Check Selected User's Access" (or double-click the user).
    2. Instant Tenant Index: Click "Pre-Index All Shared Mailboxes" once to cache all shared mailbox permissions in memory,
       allowing instant 0ms switching as you click through different users in the list.

.PARAMETER ExportCsvPath
    Directory path where exported CSV reports are saved. Defaults to .\Reports inside the tool folder.

.EXAMPLE
    .\SharedMailboxAccessViewer.ps1
    Launches the interactive WPF GUI application and prompts for Exchange Online sign-in when requested.

.EXAMPLE
    .\SharedMailboxAccessViewer.ps1 -ExportCsvPath "C:\Reports\SharedMailboxes"
    Launches the WPF GUI application with a custom default CSV export directory.

.NOTES
    Author:  James Pratt
    Version: 1.0
    Date:    28/10/2026
    Requires: PowerShell 5.1 or PowerShell 7+, ExchangeOnlineManagement module
    Logging: Writes diagnostic logs to C:\Logs\SharedMailboxAccessViewer\
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ExportCsvPath = "$PSScriptRoot\Reports"
)

# -------------------------------------------------------------------------
# STA Threading Enforcement (Required for WPF GUI)
# -------------------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$scriptPath`""
    exit
}

# -------------------------------------------------------------------------
# Load Required Assemblies & Hide Console Window (Single GUI Window Only)
# -------------------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

try {
    try {
        Add-Type -Name Win32Console -Namespace Win32Helpers -MemberDefinition @"
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@ -ErrorAction Stop
    } catch {}
    $consoleHwnd = [Win32Helpers.Win32Console]::GetConsoleWindow()
    if ($consoleHwnd -ne [IntPtr]::Zero) {
        [void][Win32Helpers.Win32Console]::ShowWindow($consoleHwnd, 0) # 0 = SW_HIDE
    }
} catch {
    Write-Verbose "Console window handle hide skipped."
}

# -------------------------------------------------------------------------
# Ensure Reports & Log Directories Exist
# -------------------------------------------------------------------------
$global:ReportsDirectory = if ($ExportCsvPath) { $ExportCsvPath } else { Join-Path $PSScriptRoot "Reports" }
if (-not (Test-Path $global:ReportsDirectory)) {
    New-Item -ItemType Directory -Path $global:ReportsDirectory -Force | Out-Null
}

$global:LogDirectory = "C:\Logs\SharedMailboxAccessViewer"
try {
    if (-not (Test-Path $global:LogDirectory)) {
        New-Item -ItemType Directory -Path $global:LogDirectory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $global:LogDirectory = Join-Path $env:LOCALAPPDATA "Logs\SharedMailboxAccessViewer"
    if (-not (Test-Path $global:LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $global:LogDirectory -Force -ErrorAction Stop | Out-Null
        } catch {}
    }
}
$global:LogFilePath = Join-Path $global:LogDirectory ("SharedMailboxAccessViewer_{0}.log" -f (Get-Date -Format "dd-MM-yy"))

# -------------------------------------------------------------------------
# Data Models for WPF DataGrids
# -------------------------------------------------------------------------
class DirectoryUserItem {
    [string]$DisplayName
    [string]$UserPrincipalName
    [string]$PrimarySmtpAddress
    [string]$Alias
    [string]$Name
    [string]$ExternalDirectoryObjectId
    [string]$DistinguishedName
    [string]$RecipientTypeDetails
    [string]$MailboxCategory
}

class MailboxAccessResultItem {
    [string]$UserDisplayName
    [string]$UserPrincipalName
    [string]$MailboxDisplayName
    [string]$MailboxEmail
    [string]$MailboxUPN
    [string]$MailboxType
    [string]$FullAccess     # "Yes" or "-"
    [string]$SendAs         # "Yes" or "-"
    [string]$SendOnBehalf   # "Yes" or "-"
    [string]$PermissionsSummary
    [string]$GALStatus      # "Visible" or "Hidden"
}

# -------------------------------------------------------------------------
# Global State & Collections
# -------------------------------------------------------------------------
$global:UserCollection   = [System.Collections.ObjectModel.ObservableCollection[DirectoryUserItem]]::new()
$global:AccessCollection = [System.Collections.ObjectModel.ObservableCollection[MailboxAccessResultItem]]::new()

$global:AllMailboxes     = @()
$global:SharedMailboxes  = @()
$global:UserMailboxes    = @()

# Cache: UserUPN (lowercase) -> List[MailboxAccessResultItem]
$global:UserAccessCache  = @{}
$global:IsFullIndexBuilt = $false
$global:IsBusy           = $false
$global:CancelRequested  = $false

# -------------------------------------------------------------------------
# UI & Dispatcher Helpers
# -------------------------------------------------------------------------
function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,
        [double]$MinWidth = 860,
        [double]$MinHeight = 640
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $fontDir = $null

    if ($PSScriptRoot) {
        $probe = $PSScriptRoot
        for ($i = 0; $i -le 5; $i++) {
            if ([string]::IsNullOrWhiteSpace($probe)) { break }
            $candidateDir = Join-Path -Path $probe -ChildPath 'assets\fonts'
            $candidateRegular = Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf'
            if (Test-Path -Path $candidateRegular -PathType Leaf) {
                $fontDir = $candidateDir
                break
            }
            $probe = Split-Path -Path $probe -Parent
        }
    }

    if (-not $fontDir) {
        $cacheFontDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        $cachedRegular = Join-Path -Path $cacheFontDir -ChildPath 'Figtree-Regular.ttf'
        if (-not (Test-Path -Path $cachedRegular -PathType Leaf)) {
            try {
                if (-not (Test-Path -Path $cacheFontDir)) {
                    New-Item -Path $cacheFontDir -ItemType Directory -Force | Out-Null
                }
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    $baseUrl = 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/'
                    foreach ($f in $fontFiles) {
                        $dest = Join-Path -Path $cacheFontDir -ChildPath $f
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            Invoke-WebRequest -Uri ($baseUrl + $f) -OutFile $dest -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                        }
                    }
                } finally {
                    $ProgressPreference = $prevProgress
                }
            } catch {
                Write-Verbose "Figtree font download skipped (offline fallback): $($_.Exception.Message)"
            }
        }
        if (Test-Path -Path $cachedRegular -PathType Leaf) {
            $fontDir = $cacheFontDir
        }
    }

    if ($fontDir) {
        try {
            $resolvedFontDir = (Resolve-Path -Path $fontDir -ErrorAction Stop).Path
            $fontUri = ([System.Uri]::new((Join-Path -Path $resolvedFontDir -ChildPath ''))).AbsoluteUri
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
        } catch {
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
        }
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Window.WindowStartupLocation = 'Manual'
    $Window.Width  = $targetWidth
    $Window.Height = $targetHeight
    $Window.MinWidth  = [Math]::Min($MinWidth, $targetWidth)
    $Window.MinHeight = [Math]::Min($MinHeight, $targetHeight)
    $Window.Left = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $Window.Top  = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Window.ResizeMode = 'CanResize'
}

function Invoke-DoEvent {
    $frame = [System.Windows.Threading.DispatcherFrame]::new()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action[System.Windows.Threading.DispatcherFrame]]{ param($f) $f.Continue = $false },
        $frame
    ) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Write-AppLog {
    param (
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $fullTimestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"
    $fileLine = "[$fullTimestamp] [$Level] $Message"

    try {
        Add-Content -Path $global:LogFilePath -Value $fileLine -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Verbose "Could not write to log file."
    }

    if ($null -ne $txtLog) {
        $txtLog.AppendText("$logLine`r`n")
        $txtLog.ScrollToEnd()
        Invoke-DoEvent
    } else {
        Write-Verbose $logLine
    }
}

function Update-AppProgress {
    param (
        [int]$Percent,
        [string]$Status
    )
    if ($null -ne $progressBar) {
        $progressBar.Value = [Math]::Min(100, [Math]::Max(0, $Percent))
    }
    if ($null -ne $lblProgress) {
        $lblProgress.Text = $Status
    }
    Invoke-DoEvent
}

function Set-BusyState {
    param ([bool]$Busy)
    $global:IsBusy = $Busy
    if ($null -ne $btnConnect)         { $btnConnect.IsEnabled = -not $Busy }
    if ($null -ne $btnReloadMailboxes) { $btnReloadMailboxes.IsEnabled = -not $Busy }
    if ($null -ne $btnPreIndexAll)     { $btnPreIndexAll.IsEnabled = -not $Busy }
    if ($null -ne $btnCheckUser)       { $btnCheckUser.IsEnabled = -not $Busy }
    if ($null -ne $btnCancel)          { $btnCancel.IsEnabled = $Busy }
    Invoke-DoEvent
}

# -------------------------------------------------------------------------
# Exchange Online Connection & Directory Loading
# -------------------------------------------------------------------------
function Connect-ExchangeService {
    param (
        [switch]$ForceNewSignIn
    )

    Write-AppLog "Checking Exchange Online PowerShell module and session..." "INFO"

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        [System.Windows.MessageBox]::Show(
            "The 'ExchangeOnlineManagement' module is not installed.`r`n`r`nRun this in PowerShell to install it:`r`nInstall-Module ExchangeOnlineManagement -Scope CurrentUser",
            "Missing Module",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        Write-AppLog "ExchangeOnlineManagement module not found." "ERROR"
        return $false
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
    } catch {}

    if ($ForceNewSignIn) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null
        } catch {
            Write-Verbose "Disconnected previous Exchange Online session."
        }
    } else {
        # Check if already connected
        try {
            $connInfo = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' })
            if ($connInfo.Count -gt 0) {
                $acct = if ($connInfo[0].UserPrincipalName) { $connInfo[0].UserPrincipalName } else { "Connected Session" }
                Update-ConnectionBadge -Connected $true -Account $acct
                Write-AppLog "Already connected to Exchange Online as $acct." "SUCCESS"
                return $true
            }
        } catch {
            Write-Verbose "No existing Exchange Online connection detected."
        }
    }

    Write-AppLog "Opening native Microsoft interactive sign-in dialog..." "INFO"
    Update-AppProgress 10 "Waiting for interactive sign-in dialog..."

    try {
        # Native interactive popup sign-in (WAM / Embedded dialog - does NOT open web browser)
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

        $connInfo = @(try { Get-ConnectionInformation -ErrorAction Stop } catch {})
        $acct = if ($connInfo.Count -gt 0 -and $connInfo[0].UserPrincipalName) { $connInfo[0].UserPrincipalName } else { "Connected" }
        Update-ConnectionBadge -Connected $true -Account $acct
        Write-AppLog "Successfully signed in to Exchange Online as $acct." "SUCCESS"
        return $true
    } catch {
        Update-ConnectionBadge -Connected $false -Account "Disconnected"
        Write-AppLog "Exchange Online sign-in failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.MessageBox]::Show(
            "Failed to sign in to Exchange Online:`r`n$($_.Exception.Message)",
            "Sign-In Failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return $false
    } finally {
        if ($null -ne $window) {
            $window.Activate() | Out-Null
            Invoke-DoEvent
        }
    }
}

function Update-ConnectionBadge {
    param (
        [bool]$Connected,
        [string]$Account
    )
    $bc = [System.Windows.Media.BrushConverter]::new()
    if ($Connected) {
        $bdgConnection.Background = $bc.ConvertFromString("#1F34D399")
        $bdgConnection.BorderBrush = $bc.ConvertFromString("#34D399")
        $lblConnectionStatus.Text = "Connected: $Account"
        $lblConnectionStatus.Foreground = $bc.ConvertFromString("#F8FAFC")
        $btnConnect.Content = "Reconnect / Switch"
    } else {
        $bdgConnection.Background = $bc.ConvertFromString("#14FFFFFF")
        $bdgConnection.BorderBrush = $bc.ConvertFromString("#26354D")
        $lblConnectionStatus.Text = "Not Signed In"
        $lblConnectionStatus.Foreground = $bc.ConvertFromString("#94A3B8")
        $btnConnect.Content = "Sign In to Exchange Online"
    }
    Invoke-DoEvent
}

function Import-TenancyMailbox {
    param (
        [switch]$ForceNewSignIn
    )
    if (-not (Connect-ExchangeService -ForceNewSignIn:$ForceNewSignIn)) { return }

    Set-BusyState $true
    try {
        $global:CancelRequested = $false
        $global:UserAccessCache.Clear()
        $global:IsFullIndexBuilt = $false
        Update-IndexBadge

        Write-AppLog "Fetching all User and Shared Mailboxes from Exchange Online..." "INFO"
        Update-AppProgress 25 "Loading users and shared mailboxes from Exchange Online..."

        $props = @(
            "DisplayName",
            "PrimarySmtpAddress",
            "UserPrincipalName",
            "Alias",
            "Name",
            "ExternalDirectoryObjectId",
            "DistinguishedName",
            "RecipientTypeDetails",
            "HiddenFromAddressListsEnabled",
            "GrantSendOnBehalfTo"
        )

        $rawMailboxes = @(Get-EXOMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -Properties $props -ResultSize Unlimited -ErrorAction Stop |
            Sort-Object DisplayName)

        $global:AllMailboxes    = $rawMailboxes
        $global:SharedMailboxes = @($rawMailboxes | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" })
        $global:UserMailboxes   = @($rawMailboxes | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" })

        Set-UserListGrid

        Write-AppLog "Loaded $($global:UserMailboxes.Count) User Mailboxes and $($global:SharedMailboxes.Count) Shared Mailboxes." "SUCCESS"
        Update-AppProgress 100 "Ready - Select a user on the left to check their Shared Mailbox access."
    } catch {
        Write-AppLog "Error loading mailboxes: $($_.Exception.Message)" "ERROR"
        Update-AppProgress 0 "Error loading mailboxes."
    } finally {
        Set-BusyState $false
    }
}

function Set-UserListGrid {
    $global:UserCollection.Clear()

    $showSharedInUserList = ($null -ne $cmbUserFilter -and $cmbUserFilter.SelectedIndex -eq 1)
    $sourceList = if ($showSharedInUserList) { $global:AllMailboxes } else { $global:UserMailboxes }

    foreach ($mb in $sourceList) {
        $item = [DirectoryUserItem]::new()
        $item.DisplayName               = if ($mb.DisplayName) { "$($mb.DisplayName)" } else { "$($mb.Name)" }
        $item.UserPrincipalName         = "$($mb.UserPrincipalName)"
        $item.PrimarySmtpAddress        = "$($mb.PrimarySmtpAddress)"
        $item.Alias                     = "$($mb.Alias)"
        $item.Name                      = "$($mb.Name)"
        $item.ExternalDirectoryObjectId = "$($mb.ExternalDirectoryObjectId)"
        $item.DistinguishedName         = "$($mb.DistinguishedName)"
        $item.RecipientTypeDetails      = "$($mb.RecipientTypeDetails)"
        $item.MailboxCategory           = if ($mb.RecipientTypeDetails -eq "SharedMailbox") { "Shared" } else { "User" }
        $global:UserCollection.Add($item)
    }

    Set-UserListFilter
}

# -------------------------------------------------------------------------
# Identity Matching Helpers
# -------------------------------------------------------------------------
function Get-UserIdentityToken {
    param ([DirectoryUserItem]$User)

    $tokens = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($val in @(
        $User.UserPrincipalName,
        $User.PrimarySmtpAddress,
        $User.DisplayName,
        $User.Alias,
        $User.Name,
        $User.ExternalDirectoryObjectId,
        $User.DistinguishedName
    )) {
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $tokens.Add($val.Trim()) | Out-Null
        }
    }
    return $tokens
}

function New-DirectoryLookupMap {
    # Maps any identifier (UPN, SMTP, DisplayName, Alias, Name, ExternalId, DN) -> DirectoryUserItem
    $map = @{}
    foreach ($mb in $global:AllMailboxes) {
        $u = [DirectoryUserItem]::new()
        $u.DisplayName               = if ($mb.DisplayName) { "$($mb.DisplayName)" } else { "$($mb.Name)" }
        $u.UserPrincipalName         = "$($mb.UserPrincipalName)"
        $u.PrimarySmtpAddress        = "$($mb.PrimarySmtpAddress)"
        $u.Alias                     = "$($mb.Alias)"
        $u.Name                      = "$($mb.Name)"
        $u.ExternalDirectoryObjectId = "$($mb.ExternalDirectoryObjectId)"
        $u.DistinguishedName         = "$($mb.DistinguishedName)"
        $u.RecipientTypeDetails      = "$($mb.RecipientTypeDetails)"

        foreach ($key in @(
            $u.UserPrincipalName,
            $u.PrimarySmtpAddress,
            $u.DisplayName,
            $u.Alias,
            $u.Name,
            $u.ExternalDirectoryObjectId,
            $u.DistinguishedName
        )) {
            if (-not [string]::IsNullOrWhiteSpace($key)) {
                $k = $key.Trim().ToLower()
                if (-not $map.ContainsKey($k)) {
                    $map[$k] = $u
                }
            }
        }
    }
    return $map
}

function Update-IndexBadge {
    $bc = [System.Windows.Media.BrushConverter]::new()
    if ($global:IsFullIndexBuilt) {
        $bdgIndexStatus.Background = $bc.ConvertFromString("#262DD4BF")
        $bdgIndexStatus.BorderBrush = $bc.ConvertFromString("#2DD4BF")
        $lblIndexStatus.Text = "Instant Index Active (Click any user)"
        $lblIndexStatus.Foreground = $bc.ConvertFromString("#5EEAD4")
    } else {
        $bdgIndexStatus.Background = $bc.ConvertFromString("#14FFFFFF")
        $bdgIndexStatus.BorderBrush = $bc.ConvertFromString("#26354D")
        $lblIndexStatus.Text = "On-Demand Mode (Select user & click Check)"
        $lblIndexStatus.Foreground = $bc.ConvertFromString("#94A3B8")
    }
    Invoke-DoEvent
}

# -------------------------------------------------------------------------
# Core Engine 1: Check Single Selected User
# -------------------------------------------------------------------------
function Get-SingleUserMailboxAccess {
    param (
        [DirectoryUserItem]$SelectedUser,
        [bool]$ForceRefresh = $false
    )

    if ($null -eq $SelectedUser) {
        Write-AppLog "Please select a user from the list first." "WARN"
        return
    }

    $includeUserMailboxes = ($null -ne $chkIncludeUserMailboxes -and [bool]$chkIncludeUserMailboxes.IsChecked)
    $cacheKey = "$($SelectedUser.UserPrincipalName.ToLower())|incUser=$includeUserMailboxes"

    # Use cached results if available and not forcing refresh
    if (-not $ForceRefresh -and $global:UserAccessCache.ContainsKey($cacheKey)) {
        Show-UserAccessResult -User $SelectedUser -Results $global:UserAccessCache[$cacheKey]
        Write-AppLog "Displayed cached access results for $($SelectedUser.DisplayName) ($($global:UserAccessCache[$cacheKey].Count) mailbox(es))." "INFO"
        return
    }

    if (-not (Connect-ExchangeService)) { return }

    Set-BusyState $true
    try {
        $global:CancelRequested = $false
        $userTokens = Get-UserIdentityToken -User $SelectedUser
        $targetPool = if ($includeUserMailboxes) { $global:AllMailboxes } else { $global:SharedMailboxes }

        if ($targetPool.Count -eq 0) {
            Write-AppLog "No mailboxes loaded yet. Loading mailboxes first..." "INFO"
            Import-TenancyMailbox
            $targetPool = if ($includeUserMailboxes) { $global:AllMailboxes } else { $global:SharedMailboxes }
        }

        Write-AppLog "Checking mailbox access for $($SelectedUser.DisplayName) <$($SelectedUser.UserPrincipalName)> across $($targetPool.Count) mailboxes..." "INFO"

        # Map of Mailbox PrimarySmtpAddress (lower) -> hashtable of flags
        $mailboxMap = @{}
        foreach ($mb in $targetPool) {
            # Skip user's own mailbox
            if ($mb.UserPrincipalName -ieq $SelectedUser.UserPrincipalName) { continue }
            $key = "$($mb.PrimarySmtpAddress)".ToLower()
            if ([string]::IsNullOrWhiteSpace($key)) { $key = "$($mb.UserPrincipalName)".ToLower() }

            $mailboxMap[$key] = @{
                Mailbox      = $mb
                FullAccess   = $false
                SendAs       = $false
                SendOnBehalf = $false
            }
        }

        # 1. Check Send on Behalf (Instant from cached GrantSendOnBehalfTo property)
        Update-AppProgress 10 "Checking Send on Behalf permissions..."
        foreach ($entry in $mailboxMap.Values) {
            $mb = $entry.Mailbox
            if ($mb.GrantSendOnBehalfTo -and $mb.GrantSendOnBehalfTo.Count -gt 0) {
                foreach ($sob in $mb.GrantSendOnBehalfTo) {
                    if ($userTokens.Contains("$sob".Trim())) {
                        $entry.SendOnBehalf = $true
                        break
                    }
                }
            }
        }

        # 2. Check Send As (Fast tenant-wide query by Trustee, with per-mailbox fallback)
        Update-AppProgress 20 "Querying Send As permissions for $($SelectedUser.UserPrincipalName)..."
        $sendAsQueriedBulk = $false
        try {
            $sendAsPerms = @(Get-EXORecipientPermission -Trustee $SelectedUser.UserPrincipalName -ResultSize Unlimited -ErrorAction Stop |
                Where-Object { $_.AccessRights -contains "SendAs" -and -not $_.IsInherited })
            $sendAsQueriedBulk = $true

            # Match returned Identity to our mailbox pool
            $identityLookup = @{}
            foreach ($entry in $mailboxMap.Values) {
                $m = $entry.Mailbox
                foreach ($idVal in @($m.PrimarySmtpAddress, $m.UserPrincipalName, $m.DisplayName, $m.Name, $m.Alias, $m.ExternalDirectoryObjectId, $m.DistinguishedName)) {
                    if (-not [string]::IsNullOrWhiteSpace($idVal)) {
                        $identityLookup["$idVal".Trim().ToLower()] = $entry
                    }
                }
            }

            foreach ($sp in $sendAsPerms) {
                $idStr = "$($sp.Identity)".Trim().ToLower()
                if ($identityLookup.ContainsKey($idStr)) {
                    $identityLookup[$idStr].SendAs = $true
                }
            }
        } catch {
            Write-AppLog "Bulk SendAs query fallback triggered; will check per mailbox." "INFO"
        }

        # 3. Check Full Access (and SendAs if bulk wasn't used) across target mailboxes
        $total = $targetPool.Count
        $current = 0

        foreach ($mb in $targetPool) {
            if ($global:CancelRequested) {
                Write-AppLog "Check cancelled by user." "WARN"
                break
            }

            $current++
            if ($mb.UserPrincipalName -ieq $SelectedUser.UserPrincipalName) { continue }

            $pct = 25 + [int](($current / [Math]::Max(1, $total)) * 70)
            if ($current % 3 -eq 0 -or $current -eq $total) {
                Update-AppProgress $pct "Checking mailbox $current of $total : $($mb.DisplayName)..."
            }

            $key = "$($mb.PrimarySmtpAddress)".ToLower()
            if ([string]::IsNullOrWhiteSpace($key)) { $key = "$($mb.UserPrincipalName)".ToLower() }
            $entry = $mailboxMap[$key]
            if ($null -eq $entry) { continue }

            $mbIdentity = if ($mb.PrimarySmtpAddress) { "$($mb.PrimarySmtpAddress)" } else { "$($mb.UserPrincipalName)" }

            # FullAccess check for this specific user
            try {
                $faPerms = @(Get-EXOMailboxPermission -Identity $mbIdentity -User $SelectedUser.UserPrincipalName -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited -and -not $_.Deny })
                if ($faPerms.Count -gt 0) {
                    $entry.FullAccess = $true
                }
            } catch {
                Write-Verbose "FullAccess query skipped for $mbIdentity"
            }

            # Fallback SendAs check if bulk Trustee query failed
            if (-not $sendAsQueriedBulk) {
                try {
                    $saPerms = @(Get-EXORecipientPermission -Identity $mbIdentity -Trustee $SelectedUser.UserPrincipalName -ErrorAction Stop |
                        Where-Object { $_.AccessRights -contains "SendAs" })
                    if ($saPerms.Count -gt 0) {
                        $entry.SendAs = $true
                    }
                } catch {
                    Write-Verbose "SendAs query skipped for $mbIdentity"
                }
            }
        }

        # Build results list
        $results = [System.Collections.Generic.List[MailboxAccessResultItem]]::new()
        foreach ($entry in $mailboxMap.Values) {
            if ($entry.FullAccess -or $entry.SendAs -or $entry.SendOnBehalf) {
                $mb = $entry.Mailbox
                $permParts = [System.Collections.Generic.List[string]]::new()
                if ($entry.FullAccess)   { $permParts.Add("FullAccess") }
                if ($entry.SendAs)       { $permParts.Add("SendAs") }
                if ($entry.SendOnBehalf) { $permParts.Add("SendOnBehalf") }

                $ri = [MailboxAccessResultItem]::new()
                $ri.UserDisplayName    = $SelectedUser.DisplayName
                $ri.UserPrincipalName  = $SelectedUser.UserPrincipalName
                $ri.MailboxDisplayName = if ($mb.DisplayName) { "$($mb.DisplayName)" } else { "$($mb.Name)" }
                $ri.MailboxEmail       = "$($mb.PrimarySmtpAddress)"
                $ri.MailboxUPN         = "$($mb.UserPrincipalName)"
                $ri.MailboxType        = if ($mb.RecipientTypeDetails -eq "SharedMailbox") { "Shared Mailbox" } else { "User Mailbox" }
                $ri.FullAccess         = if ($entry.FullAccess)   { "Yes" } else { "-" }
                $ri.SendAs             = if ($entry.SendAs)       { "Yes" } else { "-" }
                $ri.SendOnBehalf       = if ($entry.SendOnBehalf) { "Yes" } else { "-" }
                $ri.PermissionsSummary = ($permParts -join ", ")
                $ri.GALStatus          = if ([bool]$mb.HiddenFromAddressListsEnabled) { "Hidden" } else { "Visible" }
                $results.Add($ri)
            }
        }

        $sortedResults = @($results | Sort-Object MailboxDisplayName)
        $global:UserAccessCache[$cacheKey] = $sortedResults

        Show-UserAccessResult -User $SelectedUser -Results $sortedResults
        Update-AppProgress 100 "Completed check for $($SelectedUser.DisplayName): Found $($sortedResults.Count) mailbox(es)."
        Write-AppLog "Found $($sortedResults.Count) mailbox(es) accessible by $($SelectedUser.DisplayName)." "SUCCESS"
    } finally {
        Set-BusyState $false
    }
}

# -------------------------------------------------------------------------
# Core Engine 2: Pre-Index All Shared Mailboxes (0ms Instant User Switching)
# -------------------------------------------------------------------------
function Initialize-SharedMailboxIndex {
    if (-not (Connect-ExchangeService)) { return }

    if ($global:AllMailboxes.Count -eq 0) {
        Import-TenancyMailbox
    }

    $includeUserMailboxes = ($null -ne $chkIncludeUserMailboxes -and [bool]$chkIncludeUserMailboxes.IsChecked)
    $targetPool = if ($includeUserMailboxes) { $global:AllMailboxes } else { $global:SharedMailboxes }

    if ($targetPool.Count -eq 0) {
        Write-AppLog "No Shared Mailboxes found in tenant." "WARN"
        return
    }

    Set-BusyState $true
    try {
        $global:CancelRequested = $false
        $global:UserAccessCache.Clear()

        Write-AppLog "Indexing permissions across $($targetPool.Count) mailboxes for instant user switching..." "INFO"
        $userLookup = New-DirectoryLookupMap

        # Per-user accumulator: userUPNLower -> @{ mbKey -> @{ Mailbox; FullAccess; SendAs; SendOnBehalf } }
        $perUserMap = @{}

        $total = $targetPool.Count
        $i = 0

        foreach ($mb in $targetPool) {
            if ($global:CancelRequested) {
                Write-AppLog "Indexing cancelled by user." "WARN"
                return
            }

            $i++
            $pct = [int](($i / [Math]::Max(1, $total)) * 95)
            if ($i % 2 -eq 0 -or $i -eq $total) {
                Update-AppProgress $pct "Indexing mailbox $i of $total : $($mb.DisplayName)..."
            }

            $mbIdentity = if ($mb.PrimarySmtpAddress) { "$($mb.PrimarySmtpAddress)" } else { "$($mb.UserPrincipalName)" }
            $mbKey = $mbIdentity.ToLower()

            # 1. FullAccess
            try {
                $faList = @(Get-EXOMailboxPermission -Identity $mbIdentity -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited -and -not $_.Deny })
                foreach ($p in $faList) {
                    $trustee = "$($p.User)".Trim()
                    if ($trustee -like "NT AUTHORITY\*" -or $trustee -like "S-1-5-*") { continue }
                    $matchedUser = $userLookup[$trustee.ToLower()]
                    if ($null -ne $matchedUser -and $matchedUser.UserPrincipalName -ine $mb.UserPrincipalName) {
                        $uKey = $matchedUser.UserPrincipalName.ToLower()
                        if (-not $perUserMap.ContainsKey($uKey)) { $perUserMap[$uKey] = @{} }
                        if (-not $perUserMap[$uKey].ContainsKey($mbKey)) {
                            $perUserMap[$uKey][$mbKey] = @{ User = $matchedUser; Mailbox = $mb; FullAccess = $false; SendAs = $false; SendOnBehalf = $false }
                        }
                        $perUserMap[$uKey][$mbKey].FullAccess = $true
                    }
                }
            } catch {
                Write-Verbose "FullAccess index skipped for $mbIdentity"
            }

            # 2. SendAs
            try {
                $saList = @(Get-EXORecipientPermission -Identity $mbIdentity -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains "SendAs" -and -not $_.IsInherited })
                foreach ($p in $saList) {
                    $trustee = "$($p.Trustee)".Trim()
                    if ($trustee -like "NT AUTHORITY\*" -or $trustee -like "S-1-5-*") { continue }
                    $matchedUser = $userLookup[$trustee.ToLower()]
                    if ($null -ne $matchedUser -and $matchedUser.UserPrincipalName -ine $mb.UserPrincipalName) {
                        $uKey = $matchedUser.UserPrincipalName.ToLower()
                        if (-not $perUserMap.ContainsKey($uKey)) { $perUserMap[$uKey] = @{} }
                        if (-not $perUserMap[$uKey].ContainsKey($mbKey)) {
                            $perUserMap[$uKey][$mbKey] = @{ User = $matchedUser; Mailbox = $mb; FullAccess = $false; SendAs = $false; SendOnBehalf = $false }
                        }
                        $perUserMap[$uKey][$mbKey].SendAs = $true
                    }
                }
            } catch {
                Write-Verbose "SendAs index skipped for $mbIdentity"
            }

            # 3. SendOnBehalf
            if ($mb.GrantSendOnBehalfTo -and $mb.GrantSendOnBehalfTo.Count -gt 0) {
                foreach ($sob in $mb.GrantSendOnBehalfTo) {
                    $sobStr = "$sob".Trim()
                    $matchedUser = $userLookup[$sobStr.ToLower()]
                    if ($null -ne $matchedUser -and $matchedUser.UserPrincipalName -ine $mb.UserPrincipalName) {
                        $uKey = $matchedUser.UserPrincipalName.ToLower()
                        if (-not $perUserMap.ContainsKey($uKey)) { $perUserMap[$uKey] = @{} }
                        if (-not $perUserMap[$uKey].ContainsKey($mbKey)) {
                            $perUserMap[$uKey][$mbKey] = @{ User = $matchedUser; Mailbox = $mb; FullAccess = $false; SendAs = $false; SendOnBehalf = $false }
                        }
                        $perUserMap[$uKey][$mbKey].SendOnBehalf = $true
                    }
                }
            }
        }

        # Populate global cache for every user (including 0-result users so clicking them is instant)
        foreach ($mb in $global:AllMailboxes) {
            $uKey = "$($mb.UserPrincipalName)".ToLower()
            $cacheKey = "$uKey|incUser=$includeUserMailboxes"
            $userResults = [System.Collections.Generic.List[MailboxAccessResultItem]]::new()

            if ($perUserMap.ContainsKey($uKey)) {
                foreach ($entry in $perUserMap[$uKey].Values) {
                    $targetMb = $entry.Mailbox
                    $uObj     = $entry.User
                    $permParts = [System.Collections.Generic.List[string]]::new()
                    if ($entry.FullAccess)   { $permParts.Add("FullAccess") }
                    if ($entry.SendAs)       { $permParts.Add("SendAs") }
                    if ($entry.SendOnBehalf) { $permParts.Add("SendOnBehalf") }

                    $ri = [MailboxAccessResultItem]::new()
                    $ri.UserDisplayName    = $uObj.DisplayName
                    $ri.UserPrincipalName  = $uObj.UserPrincipalName
                    $ri.MailboxDisplayName = if ($targetMb.DisplayName) { "$($targetMb.DisplayName)" } else { "$($targetMb.Name)" }
                    $ri.MailboxEmail       = "$($targetMb.PrimarySmtpAddress)"
                    $ri.MailboxUPN         = "$($targetMb.UserPrincipalName)"
                    $ri.MailboxType        = if ($targetMb.RecipientTypeDetails -eq "SharedMailbox") { "Shared Mailbox" } else { "User Mailbox" }
                    $ri.FullAccess         = if ($entry.FullAccess)   { "Yes" } else { "-" }
                    $ri.SendAs             = if ($entry.SendAs)       { "Yes" } else { "-" }
                    $ri.SendOnBehalf       = if ($entry.SendOnBehalf) { "Yes" } else { "-" }
                    $ri.PermissionsSummary = ($permParts -join ", ")
                    $ri.GALStatus          = if ([bool]$targetMb.HiddenFromAddressListsEnabled) { "Hidden" } else { "Visible" }
                    $userResults.Add($ri)
                }
            }

            $global:UserAccessCache[$cacheKey] = @($userResults | Sort-Object MailboxDisplayName)
        }

        $global:IsFullIndexBuilt = $true
        Update-IndexBadge
        Update-AppProgress 100 "Full index complete! Click any user on the left to view their mailboxes instantly."
        Write-AppLog "Indexed $($targetPool.Count) mailboxes. Instant user switching is now active." "SUCCESS"

        # If a user is currently selected, show their results immediately
        $selUser = $dgUsers.SelectedItem -as [DirectoryUserItem]
        if ($null -ne $selUser) {
            Get-SingleUserMailboxAccess -SelectedUser $selUser
        }
    } finally {
        Set-BusyState $false
    }
}

# -------------------------------------------------------------------------
# UI Display & Filtering Helpers
# -------------------------------------------------------------------------
function Show-UserAccessResult {
    param (
        [DirectoryUserItem]$User,
        [MailboxAccessResultItem[]]$Results
    )

    $global:AccessCollection.Clear()

    $faCount  = 0
    $saCount  = 0
    $sobCount = 0

    foreach ($r in $Results) {
        if ($r.FullAccess -eq "Yes")   { $faCount++ }
        if ($r.SendAs -eq "Yes")       { $saCount++ }
        if ($r.SendOnBehalf -eq "Yes") { $sobCount++ }
        $global:AccessCollection.Add($r)
    }

    $lblSelectedUserName.Text  = $User.DisplayName
    $lblSelectedUserEmail.Text = "$($User.UserPrincipalName)  |  $($User.RecipientTypeDetails)"
    $lblStatTotal.Text         = "$($Results.Count)"
    $lblStatFullAccess.Text    = "$faCount"
    $lblStatSendAs.Text        = "$saCount"
    $lblStatSendOnBehalf.Text  = "$sobCount"

    Set-AccessResultFilter
}

function Set-UserListFilter {
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dgUsers.ItemsSource)
    if ($null -eq $view) { return }

    $search = if ($null -ne $txtUserSearch) { $txtUserSearch.Text.Trim() } else { "" }

    $view.Filter = [Predicate[object]]{
        param($item)
        $u = $item -as [DirectoryUserItem]
        if ($null -eq $u) { return $false }

        if (-not [string]::IsNullOrWhiteSpace($search)) {
            $matchName  = $u.DisplayName -and $u.DisplayName.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            $matchUPN   = $u.UserPrincipalName -and $u.UserPrincipalName.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            $matchEmail = $u.PrimarySmtpAddress -and $u.PrimarySmtpAddress.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            $matchAlias = $u.Alias -and $u.Alias.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            if (-not ($matchName -or $matchUPN -or $matchEmail -or $matchAlias)) { return $false }
        }
        return $true
    }

    $visibleCount = @($view).Count
    if ($null -ne $lblUserCount) {
        $lblUserCount.Text = "$visibleCount of $($global:UserCollection.Count) users"
    }
}

function Set-AccessResultFilter {
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dgAccessResults.ItemsSource)
    if ($null -eq $view) { return }

    $search   = if ($null -ne $txtResultSearch) { $txtResultSearch.Text.Trim() } else { "" }
    $permMode = if ($null -ne $cmbPermFilter)   { $cmbPermFilter.SelectedIndex }  else { 0 }

    $view.Filter = [Predicate[object]]{
        param($item)
        $r = $item -as [MailboxAccessResultItem]
        if ($null -eq $r) { return $false }

        switch ($permMode) {
            1 { if ($r.FullAccess -ne "Yes")   { return $false } }
            2 { if ($r.SendAs -ne "Yes")       { return $false } }
            3 { if ($r.SendOnBehalf -ne "Yes") { return $false } }
        }

        if (-not [string]::IsNullOrWhiteSpace($search)) {
            $mName = $r.MailboxDisplayName -and $r.MailboxDisplayName.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            $mMail = $r.MailboxEmail -and $r.MailboxEmail.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            $mPerm = $r.PermissionsSummary -and $r.PermissionsSummary.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            if (-not ($mName -or $mMail -or $mPerm)) { return $false }
        }

        return $true
    }
}

# -------------------------------------------------------------------------
# CSV Export Functions
# -------------------------------------------------------------------------
function Export-CurrentUserReport {
    if ($global:AccessCollection.Count -eq 0) {
        [System.Windows.MessageBox]::Show("There are no mailbox access results to export for the current user.", "Nothing to Export", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $selUser = $dgUsers.SelectedItem -as [DirectoryUserItem]
    $safeUser = if ($selUser) { ($selUser.UserPrincipalName -replace '[^a-zA-Z0-9._-]', '_') } else { "User" }
    $timestamp = (Get-Date).ToString("ddMMyy-HHmmss")

    $saveDlg = [System.Windows.Forms.SaveFileDialog]::new()
    $saveDlg.InitialDirectory = $global:ReportsDirectory
    $saveDlg.FileName = "SharedMailboxAccess-$safeUser-$timestamp.csv"
    $saveDlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

    if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:AccessCollection | Export-Csv -Path $saveDlg.FileName -NoTypeInformation -Encoding UTF8
            Write-AppLog "Exported $($global:AccessCollection.Count) row(s) to $($saveDlg.FileName)" "SUCCESS"
            [System.Windows.MessageBox]::Show("Report exported successfully to:`r`n$($saveDlg.FileName)", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } catch {
            Write-AppLog "Failed to export CSV: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Export-AllCachedReport {
    if (-not $global:IsFullIndexBuilt) {
        $ans = [System.Windows.MessageBox]::Show(
            "To export a full tenant report for all users, the Shared Mailboxes need to be indexed first.`r`n`r`nWould you like to build the index now?",
            "Index Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($ans -eq [System.Windows.MessageBoxResult]::Yes) {
            Initialize-SharedMailboxIndex
        } else {
            return
        }
    }

    if (-not $global:IsFullIndexBuilt) { return }

    $allRows = [System.Collections.Generic.List[MailboxAccessResultItem]]::new()
    foreach ($list in $global:UserAccessCache.Values) {
        foreach ($item in $list) {
            $allRows.Add($item)
        }
    }

    if ($allRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No shared mailbox delegations were found across the tenant.", "Empty Report", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $timestamp = (Get-Date).ToString("ddMMyy-HHmmss")
    $saveDlg = [System.Windows.Forms.SaveFileDialog]::new()
    $saveDlg.InitialDirectory = $global:ReportsDirectory
    $saveDlg.FileName = "Tenant-SharedMailboxAccess-AllUsers-$timestamp.csv"
    $saveDlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

    if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $allRows | Sort-Object UserDisplayName, MailboxDisplayName | Export-Csv -Path $saveDlg.FileName -NoTypeInformation -Encoding UTF8
            Write-AppLog "Exported full tenant report ($($allRows.Count) delegation rows) to $($saveDlg.FileName)" "SUCCESS"
            [System.Windows.MessageBox]::Show("Full tenant report exported ($($allRows.Count) delegations):`r`n$($saveDlg.FileName)", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } catch {
            Write-AppLog "Failed to export full tenant CSV: $($_.Exception.Message)" "ERROR"
        }
    }
}

# -------------------------------------------------------------------------
# WPF XAML User Interface (Obsidian Teal Standard)
# -------------------------------------------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft 365 Shared Mailbox Access Viewer v1.0"
        Height="820" Width="1340" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual" Foreground="#F8FAFC" FontSize="12.5">
    <Window.Background>
        <RadialGradientBrush Center="0.5,0.0" GradientOrigin="0.5,0.0" RadiusX="0.95" RadiusY="0.85">
            <GradientStop Color="#0F1F3D" Offset="0.0"/>
            <GradientStop Color="#091122" Offset="0.55"/>
            <GradientStop Color="#05070C" Offset="1.0"/>
        </RadialGradientBrush>
    </Window.Background>

    <Window.Resources>
        <!-- Secondary Button -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#14FFFFFF"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#24FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#0DFFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary CTA (Teal) -->
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#2DD4BF"/>
            <Setter Property="Foreground" Value="#042F2E"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5EEAD4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#14B8A6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#1A2B44"/>
                                <Setter Property="Foreground" Value="#64748B"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Teal Outline Accent Button -->
        <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#1A2DD4BF"/>
            <Setter Property="Foreground" Value="#5EEAD4"/>
            <Setter Property="BorderBrush" Value="#2DD4BF"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>

        <!-- Danger Button -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#26EF4444"/>
            <Setter Property="Foreground" Value="#FCA5A5"/>
            <Setter Property="BorderBrush" Value="#66EF4444"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3CEF4444"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}"/>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#99070D1A"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="CaretBrush" Value="#2DD4BF"/>
            <Setter Property="SelectionBrush" Value="#2DD4BF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Border x:Name="Box" Width="16" Height="16" CornerRadius="4"
                                    Background="#99070D1A" BorderBrush="#26354D" BorderThickness="1"
                                    Margin="0,0,7,0" VerticalAlignment="Center">
                                <Path x:Name="CheckMark" Data="M 3,7.5 L 6.2,10.5 L 11.5,4"
                                      Stroke="#042F2E" StrokeThickness="2" Visibility="Collapsed"
                                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"/>
                            </Border>
                            <ContentPresenter VerticalAlignment="Center"/>
                        </StackPanel>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box" Property="Background" Value="#2DD4BF"/>
                                <Setter TargetName="Box" Property="BorderBrush" Value="#2DD4BF"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Box" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark ComboBox & ComboBoxItem -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="#1E3A5F"/>
                                <Setter Property="Foreground" Value="#F8FAFC"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="ItemBorder" Property="Background" Value="#262DD4BF"/>
                                <Setter Property="Foreground" Value="#F8FAFC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#99070D1A"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Focusable="False"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border x:Name="ComboBorder" Background="#99070D1A" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="26"/>
                                                </Grid.ColumnDefinitions>
                                                <Path Grid.Column="1" Data="M 0,0 L 4,4 L 8,0 Z" Fill="#94A3B8" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Grid>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="ComboBorder" Property="BorderBrush" Value="#2DD4BF"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              Margin="10,4,28,4" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid MinWidth="{TemplateBinding ActualWidth}" MaxHeight="260" SnapsToDevicePixels="True">
                                    <Border Background="#0D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Margin="0,2,0,0">
                                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark DataGrid -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#091122"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="#091122"/>
            <Setter Property="AlternatingRowBackground" Value="#0C172B"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#1A2B44"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#0B1528"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="#091122"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#142DD4BF"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1E3A5F"/>
                    <Setter Property="Foreground" Value="#F8FAFC"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Foreground" Value="{Binding Foreground, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="Foreground" Value="#F8FAFC"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#050A14"/>
            <Setter Property="Width" Value="10"/>
            <Setter Property="MinWidth" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border Background="#26354D" CornerRadius="4" Margin="2"/>
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

        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#26354D" BorderThickness="1,1,1,1" CornerRadius="6,6,0,0">
                                <TabPanel IsItemsHost="True" Margin="6,0"/>
                            </Border>
                            <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#E2E8F0"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Foreground" Value="#F8FAFC"/>
                                <Setter TargetName="TabBorder" Property="Background" Value="#142DD4BF"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
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
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- TOP HEADER BAR -->
        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="20,13">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock Text="Shared Mailbox Access Viewer v1.0" Foreground="#F8FAFC" FontSize="19" FontWeight="Bold"/>
                    <TextBlock Text="Inspect Exchange Online Shared Mailbox Full Access, Send As, and Send on Behalf delegations."
                               Foreground="#94A3B8" FontSize="12" Margin="0,2,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Border x:Name="bdgConnection" Background="#14FFFFFF" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="11,6" Margin="0,0,10,0" VerticalAlignment="Center">
                        <TextBlock x:Name="lblConnectionStatus" Text="Not Signed In" Foreground="#94A3B8" FontSize="12" FontWeight="SemiBold"/>
                    </Border>
                    <Button x:Name="btnConnect" Content="Sign In to Exchange Online" Style="{StaticResource PrimaryButton}" Height="34" Margin="0,0,8,0"/>
                    <Button x:Name="btnReloadMailboxes" Content="Reload Users" Height="34" Margin="0,0,8,0"/>
                    <Button x:Name="btnPreIndexAll" Content="Pre-Index All Shared Mailboxes" Style="{StaticResource AccentButton}" Height="34" ToolTip="Scans all Shared Mailboxes once so clicking any user displays their access in 0ms"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN SPLIT CONTENT -->
        <Grid Grid.Row="1" Margin="16,12,16,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="420" MinWidth="300"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT PANE: USER DIRECTORY LIST -->
            <Border Grid.Column="0" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="8" Padding="12">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Section Title -->
                    <Grid Grid.Row="0" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="1. Select User" FontSize="14.5" FontWeight="Bold" Foreground="#F8FAFC"/>
                        <TextBlock x:Name="lblUserCount" Grid.Column="1" Text="0 users loaded" FontSize="12" FontWeight="SemiBold" Foreground="#2DD4BF" VerticalAlignment="Center"/>
                    </Grid>

                    <!-- Search Box -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Search:" VerticalAlignment="Center" Foreground="#94A3B8" FontWeight="SemiBold" Margin="0,0,8,0" FontSize="12"/>
                        <TextBox x:Name="txtUserSearch" Grid.Column="1" Height="32" VerticalContentAlignment="Center"
                                 ToolTip="Type a user's name, email, or alias to filter"/>
                        <Button x:Name="btnClearUserSearch" Grid.Column="2" Content="X" Width="32" Height="32" Padding="0" Margin="6,0,0,0" ToolTip="Clear search"/>
                    </Grid>

                    <!-- Mailbox Type Filter -->
                    <Grid Grid.Row="2" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Show:" VerticalAlignment="Center" Foreground="#94A3B8" FontWeight="SemiBold" Margin="0,0,10,0" FontSize="12"/>
                        <ComboBox x:Name="cmbUserFilter" Grid.Column="1" Height="30" SelectedIndex="0" VerticalContentAlignment="Center">
                            <ComboBoxItem Content="User Mailboxes Only (Standard Users)"/>
                            <ComboBoxItem Content="All Mailboxes (Users + Shared Mailboxes)"/>
                        </ComboBox>
                    </Grid>

                    <!-- Users DataGrid -->
                    <DataGrid x:Name="dgUsers" Grid.Row="3" AutoGenerateColumns="False" CanUserAddRows="False"
                              SelectionMode="Single" IsReadOnly="True" RowHeight="30" FontSize="12">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="1.2*"/>
                            <DataGridTextColumn Header="Email / UPN" Binding="{Binding PrimarySmtpAddress}" Width="1.5*"/>
                            <DataGridTextColumn Header="Type" Binding="{Binding MailboxCategory}" Width="65"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Bottom Left Action Button -->
                    <Grid Grid.Row="4" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="btnCheckUser" Grid.Column="0" Content="Check Selected User's Access" Style="{StaticResource PrimaryButton}" Height="36"/>
                        <Button x:Name="btnCancel" Grid.Column="1" Content="Stop" Style="{StaticResource DangerButton}" Height="36" MinWidth="72" Margin="8,0,0,0" IsEnabled="False"/>
                    </Grid>
                </Grid>
            </Border>

            <!-- SPLITTER -->
            <GridSplitter Grid.Column="1" Width="12" HorizontalAlignment="Stretch" Background="Transparent"/>

            <!-- RIGHT PANE: SELECTED USER'S MAILBOX ACCESS -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- User Summary & KPI Strip (Flat container -> single-level KPI tiles) -->
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel VerticalAlignment="Center" Margin="0,0,12,0">
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                            <TextBlock x:Name="lblSelectedUserName" Text="No User Selected" FontSize="17" FontWeight="Bold" Foreground="#F8FAFC"/>
                            <Border x:Name="bdgIndexStatus" Background="#14FFFFFF" BorderBrush="#26354D" BorderThickness="1" CornerRadius="4" Padding="8,2" Margin="10,0,0,0" VerticalAlignment="Center">
                                <TextBlock x:Name="lblIndexStatus" Text="On-Demand Mode (Select user &amp; click Check)" FontSize="11" FontWeight="SemiBold" Foreground="#94A3B8"/>
                            </Border>
                        </StackPanel>
                        <TextBlock x:Name="lblSelectedUserEmail" Text="Sign in and select a user on the left to inspect their Shared Mailbox permissions." FontSize="12" Foreground="#94A3B8" TextTrimming="CharacterEllipsis"/>
                    </StackPanel>

                    <!-- KPI Counters -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Border Background="#0D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="12,7" Margin="0,0,6,0" MinWidth="84">
                            <StackPanel HorizontalAlignment="Center">
                                <TextBlock x:Name="lblStatTotal" Text="0" FontSize="18" FontWeight="Bold" Foreground="#2DD4BF" FontFamily="Cascadia Mono, Consolas" HorizontalAlignment="Center"/>
                                <TextBlock Text="Mailboxes" FontSize="10.5" FontWeight="SemiBold" Foreground="#94A3B8" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Background="#0D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="12,7" Margin="0,0,6,0" MinWidth="84">
                            <StackPanel HorizontalAlignment="Center">
                                <TextBlock x:Name="lblStatFullAccess" Text="0" FontSize="18" FontWeight="Bold" Foreground="#34D399" FontFamily="Cascadia Mono, Consolas" HorizontalAlignment="Center"/>
                                <TextBlock Text="Full Access" FontSize="10.5" FontWeight="SemiBold" Foreground="#94A3B8" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Background="#0D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="12,7" Margin="0,0,6,0" MinWidth="84">
                            <StackPanel HorizontalAlignment="Center">
                                <TextBlock x:Name="lblStatSendAs" Text="0" FontSize="18" FontWeight="Bold" Foreground="#FBBF24" FontFamily="Cascadia Mono, Consolas" HorizontalAlignment="Center"/>
                                <TextBlock Text="Send As" FontSize="10.5" FontWeight="SemiBold" Foreground="#94A3B8" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <Border Background="#0D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="12,7" MinWidth="92">
                            <StackPanel HorizontalAlignment="Center">
                                <TextBlock x:Name="lblStatSendOnBehalf" Text="0" FontSize="18" FontWeight="Bold" Foreground="#CBD5E1" FontFamily="Cascadia Mono, Consolas" HorizontalAlignment="Center"/>
                                <TextBlock Text="Send On Behalf" FontSize="10.5" FontWeight="SemiBold" Foreground="#94A3B8" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Grid>

                <!-- Tabbed Results & Activity Log -->
                <TabControl Grid.Row="1">
                    <TabItem Header="Permissions Grid">
                        <Grid Margin="0,8,0,0">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>

                            <!-- Filter Bar -->
                            <Border Grid.Row="0" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="0,0,0,8">
                                <WrapPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,2,14,2">
                                        <TextBlock Text="Filter Mailboxes:" VerticalAlignment="Center" Foreground="#94A3B8" FontWeight="SemiBold" Margin="0,0,6,0" FontSize="12"/>
                                        <TextBox x:Name="txtResultSearch" Width="170" Height="30" VerticalContentAlignment="Center"/>
                                    </StackPanel>

                                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,2,14,2">
                                        <TextBlock Text="Permission:" VerticalAlignment="Center" Foreground="#94A3B8" FontWeight="SemiBold" Margin="0,0,6,0" FontSize="12"/>
                                        <ComboBox x:Name="cmbPermFilter" Width="165" Height="30" SelectedIndex="0" VerticalContentAlignment="Center">
                                            <ComboBoxItem Content="All Permissions"/>
                                            <ComboBoxItem Content="Full Access Only"/>
                                            <ComboBoxItem Content="Send As Only"/>
                                            <ComboBoxItem Content="Send On Behalf Only"/>
                                        </ComboBox>
                                    </StackPanel>

                                    <CheckBox x:Name="chkIncludeUserMailboxes" VerticalAlignment="Center" Margin="0,2,0,2"
                                              ToolTip="Also check if the user has delegated access to other regular User Mailboxes">
                                        <TextBlock Text="Include Delegated User Mailboxes" FontSize="12" Foreground="#CBD5E1" FontWeight="SemiBold"/>
                                    </CheckBox>
                                </WrapPanel>
                            </Border>

                            <!-- Results DataGrid -->
                            <DataGrid x:Name="dgAccessResults" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False"
                                      IsReadOnly="True" RowHeight="32" FontSize="12">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Shared Mailbox Name" Binding="{Binding MailboxDisplayName}" Width="1.6*"/>
                                    <DataGridTextColumn Header="Mailbox Email" Binding="{Binding MailboxEmail}" Width="1.8*"/>
                                    <DataGridTextColumn Header="Type" Binding="{Binding MailboxType}" Width="115"/>
                                    <DataGridTextColumn Header="Full Access" Binding="{Binding FullAccess}" Width="85"/>
                                    <DataGridTextColumn Header="Send As" Binding="{Binding SendAs}" Width="80"/>
                                    <DataGridTextColumn Header="Send On Behalf" Binding="{Binding SendOnBehalf}" Width="105"/>
                                    <DataGridTextColumn Header="Permissions Summary" Binding="{Binding PermissionsSummary}" Width="1.4*"/>
                                    <DataGridTextColumn Header="GAL Status" Binding="{Binding GALStatus}" Width="85"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </TabItem>

                    <TabItem Header="Activity Log">
                        <Border Margin="0,8,0,0" Background="#CC050A14" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="10">
                            <TextBox x:Name="txtLog" Background="Transparent" Foreground="#CBD5E1" BorderThickness="0" Padding="0"
                                     FontFamily="Cascadia Mono, Consolas" FontSize="11.5" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
                        </Border>
                    </TabItem>
                </TabControl>
            </Grid>
        </Grid>

        <!-- SLIM PROGRESS BAR -->
        <ProgressBar x:Name="progressBar" Grid.Row="2" Height="3" Minimum="0" Maximum="100" Value="0"
                     Background="#0B1528" Foreground="#2DD4BF" BorderThickness="0"/>

        <!-- SINGLE-ROW FOOTER -->
        <Border Grid.Row="3" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="lblProgress" Grid.Column="0" Text="Ready. Click 'Sign In to Exchange Online' to load users."
                           Foreground="#94A3B8" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"
                           TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="btnExportUserCsv" Content="Export User CSV" Height="34" Margin="0,0,10,0"/>
                    <Button x:Name="btnExportAllCsv" Content="Export All Users CSV" Style="{StaticResource SuccessButton}" Height="34"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# -------------------------------------------------------------------------
# Initialize WPF Window & Bind Controls
# -------------------------------------------------------------------------
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $window

$bdgConnection          = $window.FindName("bdgConnection")
$lblConnectionStatus    = $window.FindName("lblConnectionStatus")
$btnConnect             = $window.FindName("btnConnect")
$btnReloadMailboxes     = $window.FindName("btnReloadMailboxes")
$btnPreIndexAll         = $window.FindName("btnPreIndexAll")

$lblUserCount           = $window.FindName("lblUserCount")
$txtUserSearch          = $window.FindName("txtUserSearch")
$btnClearUserSearch     = $window.FindName("btnClearUserSearch")
$cmbUserFilter          = $window.FindName("cmbUserFilter")
$dgUsers                = $window.FindName("dgUsers")
$btnCheckUser           = $window.FindName("btnCheckUser")
$btnCancel              = $window.FindName("btnCancel")

$lblSelectedUserName    = $window.FindName("lblSelectedUserName")
$lblSelectedUserEmail   = $window.FindName("lblSelectedUserEmail")
$bdgIndexStatus         = $window.FindName("bdgIndexStatus")
$lblIndexStatus         = $window.FindName("lblIndexStatus")
$lblStatTotal           = $window.FindName("lblStatTotal")
$lblStatFullAccess      = $window.FindName("lblStatFullAccess")
$lblStatSendAs          = $window.FindName("lblStatSendAs")
$lblStatSendOnBehalf    = $window.FindName("lblStatSendOnBehalf")

$txtResultSearch        = $window.FindName("txtResultSearch")
$cmbPermFilter          = $window.FindName("cmbPermFilter")
$chkIncludeUserMailboxes= $window.FindName("chkIncludeUserMailboxes")
$btnExportUserCsv       = $window.FindName("btnExportUserCsv")
$btnExportAllCsv        = $window.FindName("btnExportAllCsv")
$dgAccessResults        = $window.FindName("dgAccessResults")

$lblProgress            = $window.FindName("lblProgress")
$progressBar            = $window.FindName("progressBar")
$txtLog                 = $window.FindName("txtLog")

# Bind ObservableCollections
$dgUsers.ItemsSource         = $global:UserCollection
$dgAccessResults.ItemsSource = $global:AccessCollection

# -------------------------------------------------------------------------
# Wire Event Handlers
# -------------------------------------------------------------------------
$btnConnect.Add_Click({
    $isAlreadyConn = ($btnConnect.Content -like "*Reconnect*")
    Import-TenancyMailbox -ForceNewSignIn:$isAlreadyConn
})

$btnReloadMailboxes.Add_Click({
    Import-TenancyMailbox
})

$btnPreIndexAll.Add_Click({
    Initialize-SharedMailboxIndex
})

$btnCheckUser.Add_Click({
    $sel = $dgUsers.SelectedItem -as [DirectoryUserItem]
    Get-SingleUserMailboxAccess -SelectedUser $sel -ForceRefresh $true
})

$btnCancel.Add_Click({
    $global:CancelRequested = $true
    Write-AppLog "Cancellation requested..." "WARN"
})

$txtUserSearch.Add_TextChanged({
    Set-UserListFilter
})

$btnClearUserSearch.Add_Click({
    $txtUserSearch.Text = ""
})

$cmbUserFilter.Add_SelectionChanged({
    if ($global:AllMailboxes.Count -gt 0) {
        Set-UserListGrid
    }
})

# When a user is selected in the left DataGrid:
# - If full index or cached result exists for this user, display immediately (0ms).
# - Otherwise update the header and prompt to click "Check Selected User's Access" (or double-click).
$dgUsers.Add_SelectionChanged({
    $sel = $dgUsers.SelectedItem -as [DirectoryUserItem]
    if ($null -eq $sel) { return }

    $includeUserMailboxes = [bool]$chkIncludeUserMailboxes.IsChecked
    $cacheKey = "$($sel.UserPrincipalName.ToLower())|incUser=$includeUserMailboxes"

    if ($global:IsFullIndexBuilt -or $global:UserAccessCache.ContainsKey($cacheKey)) {
        Get-SingleUserMailboxAccess -SelectedUser $sel -ForceRefresh $false
    } else {
        $global:AccessCollection.Clear()
        $lblSelectedUserName.Text  = $sel.DisplayName
        $lblSelectedUserEmail.Text = "$($sel.UserPrincipalName)  |  Click 'Check Selected User's Access' (or double-click user) to scan"
        $lblStatTotal.Text         = "-"
        $lblStatFullAccess.Text    = "-"
        $lblStatSendAs.Text        = "-"
        $lblStatSendOnBehalf.Text  = "-"
    }
})

# Double-clicking a user in the left DataGrid immediately checks their access
$dgUsers.Add_MouseDoubleClick({
    if (-not $global:IsBusy) {
        $sel = $dgUsers.SelectedItem -as [DirectoryUserItem]
        if ($null -ne $sel) {
            Get-SingleUserMailboxAccess -SelectedUser $sel -ForceRefresh $false
        }
    }
})

$txtResultSearch.Add_TextChanged({
    Set-AccessResultFilter
})

$cmbPermFilter.Add_SelectionChanged({
    Set-AccessResultFilter
})

$chkIncludeUserMailboxes.Add_Click({
    $global:IsFullIndexBuilt = $false
    Update-IndexBadge
    $sel = $dgUsers.SelectedItem -as [DirectoryUserItem]
    if ($null -ne $sel) {
        $includeUserMailboxes = [bool]$chkIncludeUserMailboxes.IsChecked
        $cacheKey = "$($sel.UserPrincipalName.ToLower())|incUser=$includeUserMailboxes"
        if ($global:UserAccessCache.ContainsKey($cacheKey)) {
            Get-SingleUserMailboxAccess -SelectedUser $sel -ForceRefresh $false
        }
    }
})

$btnExportUserCsv.Add_Click({
    Export-CurrentUserReport
})

$btnExportAllCsv.Add_Click({
    Export-AllCachedReport
})

# Initial Startup Log
$window.Add_Loaded({
    Write-AppLog "Shared Mailbox Access Viewer initialized." "INFO"
    Write-AppLog "Click 'Sign In to Exchange Online' to authenticate and load all users & shared mailboxes." "INFO"
})

[void]$window.ShowDialog()
