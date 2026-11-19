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
    if ($Connected) {
        $bdgConnection.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#065F46")
        $bdgConnection.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
        $lblConnectionStatus.Text = "Connected: $Account"
        $lblConnectionStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ECFDF5")
        $btnConnect.Content = "Reconnect / Switch"
    } else {
        $bdgConnection.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1E293B")
        $bdgConnection.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#475569")
        $lblConnectionStatus.Text = "Not Signed In"
        $lblConnectionStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#94A3B8")
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
    if ($global:IsFullIndexBuilt) {
        $bdgIndexStatus.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#DBEAFE")
        $bdgIndexStatus.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $lblIndexStatus.Text = "Instant Index Active (Click any user)"
        $lblIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1D4ED8")
    } else {
        $bdgIndexStatus.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F1F5F9")
        $bdgIndexStatus.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#CBD5E1")
        $lblIndexStatus.Text = "On-Demand Mode (Select user & click Check)"
        $lblIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#475569")
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
# WPF XAML User Interface
# -------------------------------------------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft 365 Shared Mailbox Access Viewer"
        Height="820" Width="1340" MinHeight="640" MinWidth="1020"
        WindowStartupLocation="CenterScreen"
        Background="#F1F5F9"
        FontFamily="Segoe UI">

    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
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
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#7C3AED"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#0D9488"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="14"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- TOP HEADER BAR -->
        <Border Grid.Row="0" Background="#0F172A" Padding="18,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Shared Mailbox Access Viewer" Foreground="#FFFFFF" FontSize="19" FontWeight="Bold"/>
                        <Border Background="#1E293B" CornerRadius="4" Margin="10,2,0,0" Padding="8,2" VerticalAlignment="Center">
                            <TextBlock Text="Exchange Online" Foreground="#38BDF8" FontSize="11" FontWeight="Bold"/>
                        </Border>
                    </StackPanel>
                    <TextBlock Text="Sign in once to Exchange Online, select any user, and view all Shared Mailboxes they have Full Access, Send As, or Send on Behalf permissions for."
                               Foreground="#94A3B8" FontSize="12" Margin="0,3,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Border x:Name="bdgConnection" Background="#1E293B" BorderBrush="#475569" BorderThickness="1" CornerRadius="6" Padding="10,6" Margin="0,0,10,0" VerticalAlignment="Center">
                        <TextBlock x:Name="lblConnectionStatus" Text="Not Signed In" Foreground="#94A3B8" FontSize="12" FontWeight="SemiBold"/>
                    </Border>
                    <Button x:Name="btnConnect" Content="Sign In to Exchange Online" Style="{StaticResource PrimaryButton}" Height="34" Margin="0,0,8,0"/>
                    <Button x:Name="btnReloadMailboxes" Content="Reload Users" Height="34" Margin="0,0,8,0"/>
                    <Button x:Name="btnPreIndexAll" Content="Pre-Index All Shared Mailboxes" Style="{StaticResource AccentButton}" Height="34" ToolTip="Scans all Shared Mailboxes once so clicking any user displays their access in 0ms"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN SPLIT CONTENT -->
        <Grid Grid.Row="1" Margin="12,10,12,6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="430" MinWidth="320"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT PANE: USER DIRECTORY LIST -->
            <Border Grid.Column="0" Style="{StaticResource Card}" Padding="12">
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
                        <TextBlock Text="1. Select User" FontSize="15" FontWeight="Bold" Foreground="#0F172A"/>
                        <TextBlock x:Name="lblUserCount" Grid.Column="1" Text="0 users loaded" FontSize="12" FontWeight="SemiBold" Foreground="#2563EB" VerticalAlignment="Center"/>
                    </Grid>

                    <!-- Search Box -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Search:" VerticalAlignment="Center" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,6,0" FontSize="12"/>
                        <TextBox x:Name="txtUserSearch" Grid.Column="1" Height="30" VerticalContentAlignment="Center" Padding="6,0"
                                 ToolTip="Type a user's name, email, or alias to filter"/>
                        <Button x:Name="btnClearUserSearch" Grid.Column="2" Content="X" Width="30" Height="30" Padding="0" Margin="4,0,0,0" ToolTip="Clear search"/>
                    </Grid>

                    <!-- Mailbox Type Filter -->
                    <Grid Grid.Row="2" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Show:" VerticalAlignment="Center" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,8,0" FontSize="12"/>
                        <ComboBox x:Name="cmbUserFilter" Grid.Column="1" Height="28" SelectedIndex="0" VerticalContentAlignment="Center">
                            <ComboBoxItem Content="User Mailboxes Only (Standard Users)"/>
                            <ComboBoxItem Content="All Mailboxes (Users + Shared Mailboxes)"/>
                        </ComboBox>
                    </Grid>

                    <!-- Users DataGrid -->
                    <DataGrid x:Name="dgUsers" Grid.Row="3" AutoGenerateColumns="False" CanUserAddRows="False"
                              SelectionMode="Single" IsReadOnly="True"
                              Background="#FFFFFF" BorderBrush="#E2E8F0" BorderThickness="1"
                              RowHeaderWidth="0" GridLinesVisibility="Horizontal" HeadersVisibility="Column"
                              RowHeight="30" FontSize="12">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="1.2*"/>
                            <DataGridTextColumn Header="Email / UPN" Binding="{Binding PrimarySmtpAddress}" Width="1.5*"/>
                            <DataGridTextColumn Header="Type" Binding="{Binding MailboxCategory}" Width="60"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Bottom Left Action Button -->
                    <Grid Grid.Row="4" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="btnCheckUser" Grid.Column="0" Content="Check Selected User's Access" Style="{StaticResource PrimaryButton}" Height="36"/>
                        <Button x:Name="btnCancel" Grid.Column="1" Content="Stop" Height="36" Margin="6,0,0,0" IsEnabled="False"/>
                    </Grid>
                </Grid>
            </Border>

            <!-- SPLITTER -->
            <GridSplitter Grid.Column="1" Width="10" HorizontalAlignment="Stretch" Background="Transparent"/>

            <!-- RIGHT PANE: SELECTED USER'S MAILBOX ACCESS -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- User Summary & KPI Card -->
                <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,8" Padding="14,12">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                                <TextBlock x:Name="lblSelectedUserName" Text="No User Selected" FontSize="17" FontWeight="Bold" Foreground="#0F172A"/>
                                <Border x:Name="bdgIndexStatus" Background="#F1F5F9" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="4" Padding="8,2" Margin="12,0,0,0" VerticalAlignment="Center">
                                    <TextBlock x:Name="lblIndexStatus" Text="On-Demand Mode (Select user &amp; click Check)" FontSize="11" FontWeight="SemiBold" Foreground="#475569"/>
                                </Border>
                            </StackPanel>
                            <TextBlock x:Name="lblSelectedUserEmail" Text="Sign in and select a user on the left to inspect their Shared Mailbox permissions." FontSize="12" Foreground="#64748B"/>
                        </StackPanel>

                        <!-- KPI Counters -->
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Border Background="#EFF6FF" BorderBrush="#BFDBFE" BorderThickness="1" CornerRadius="6" Padding="12,6" Margin="0,0,8,0" MinWidth="85">
                                <StackPanel HorizontalAlignment="Center">
                                    <TextBlock x:Name="lblStatTotal" Text="0" FontSize="18" FontWeight="Bold" Foreground="#1D4ED8" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Mailboxes" FontSize="10" FontWeight="SemiBold" Foreground="#3B82F6" HorizontalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                            <Border Background="#ECFDF5" BorderBrush="#A7F3D0" BorderThickness="1" CornerRadius="6" Padding="12,6" Margin="0,0,8,0" MinWidth="85">
                                <StackPanel HorizontalAlignment="Center">
                                    <TextBlock x:Name="lblStatFullAccess" Text="0" FontSize="18" FontWeight="Bold" Foreground="#047857" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Full Access" FontSize="10" FontWeight="SemiBold" Foreground="#10B981" HorizontalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                            <Border Background="#F5F3FF" BorderBrush="#DDD6FE" BorderThickness="1" CornerRadius="6" Padding="12,6" Margin="0,0,8,0" MinWidth="85">
                                <StackPanel HorizontalAlignment="Center">
                                    <TextBlock x:Name="lblStatSendAs" Text="0" FontSize="18" FontWeight="Bold" Foreground="#6D28D9" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Send As" FontSize="10" FontWeight="SemiBold" Foreground="#8B5CF6" HorizontalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                            <Border Background="#FFFBEB" BorderBrush="#FDE68A" BorderThickness="1" CornerRadius="6" Padding="12,6" MinWidth="85">
                                <StackPanel HorizontalAlignment="Center">
                                    <TextBlock x:Name="lblStatSendOnBehalf" Text="0" FontSize="18" FontWeight="Bold" Foreground="#B45309" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Send On Behalf" FontSize="10" FontWeight="SemiBold" Foreground="#F59E0B" HorizontalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Filter & Export Bar -->
                <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,0,0,8" Padding="10,8">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,12,0">
                            <TextBlock Text="Filter Mailboxes:" VerticalAlignment="Center" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,6,0" FontSize="12"/>
                            <TextBox x:Name="txtResultSearch" Width="170" Height="28" VerticalContentAlignment="Center" Padding="6,0"/>
                        </StackPanel>

                        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,12,0">
                            <TextBlock Text="Permission:" VerticalAlignment="Center" Foreground="#64748B" FontWeight="SemiBold" Margin="0,0,6,0" FontSize="12"/>
                            <ComboBox x:Name="cmbPermFilter" Width="155" Height="28" SelectedIndex="0" VerticalContentAlignment="Center">
                                <ComboBoxItem Content="All Permissions"/>
                                <ComboBoxItem Content="Full Access Only"/>
                                <ComboBoxItem Content="Send As Only"/>
                                <ComboBoxItem Content="Send On Behalf Only"/>
                            </ComboBox>
                        </StackPanel>

                        <CheckBox x:Name="chkIncludeUserMailboxes" Grid.Column="2" VerticalAlignment="Center" Margin="0,0,10,0"
                                  ToolTip="Also check if the user has delegated access to other regular User Mailboxes">
                            <TextBlock Text="Include Delegated User Mailboxes" FontSize="12" Foreground="#334155" FontWeight="SemiBold"/>
                        </CheckBox>

                        <Button x:Name="btnExportUserCsv" Grid.Column="4" Content="Export User CSV" Margin="0,0,6,0"/>
                        <Button x:Name="btnExportAllCsv" Grid.Column="5" Content="Export All Users CSV" Style="{StaticResource SuccessButton}"/>
                    </Grid>
                </Border>

                <!-- Results DataGrid -->
                <Border Grid.Row="2" Style="{StaticResource Card}" Padding="0">
                    <DataGrid x:Name="dgAccessResults" AutoGenerateColumns="False" CanUserAddRows="False"
                              IsReadOnly="True" Background="#FFFFFF" BorderThickness="0"
                              RowHeaderWidth="0" GridLinesVisibility="Horizontal" HeadersVisibility="Column"
                              RowHeight="32" FontSize="12">
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
                </Border>
            </Grid>
        </Grid>

        <!-- BOTTOM PROGRESS & ACTIVITY LOG -->
        <Border Grid.Row="2" Background="#0F172A" Padding="14,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="85"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="260"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock x:Name="lblProgress" Text="Ready. Click 'Sign In to Exchange Online' to load users." Foreground="#E2E8F0" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <ProgressBar x:Name="progressBar" Grid.Column="1" Height="12" Minimum="0" Maximum="100" Value="0" Background="#1E293B" Foreground="#38BDF8" BorderThickness="0"/>
                </Grid>

                <TextBox x:Name="txtLog" Grid.Row="1" Background="#020617" Foreground="#94A3B8" BorderBrush="#1E293B" BorderThickness="1"
                         FontFamily="Consolas" FontSize="11" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Padding="6"/>
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
