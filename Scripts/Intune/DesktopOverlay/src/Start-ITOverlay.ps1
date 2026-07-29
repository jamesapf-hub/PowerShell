<#
.SYNOPSIS
    Enterprise Desktop Overlay - IT Support Helpdesk Info
.DESCRIPTION
    Renders a non-interactive, click-through desktop overlay displaying
    IT support contact details and system diagnostic information.
    Attaches to the Desktop Background layer (behind application windows).
#>

param(
    [string]$Title = "IT SUPPORT HELPDESK",
    [string]$SupportPhone = "0800 123 4567",
    [string]$SupportEmail = "support@company.com",
    [string]$SupportHours = "Mon-Fri: 08:30 - 17:00",
    [ValidateSet("PhoneFirst", "EmailFirst")][string]$FieldOrder = "EmailFirst",
    [ValidateSet("BottomRight", "TopRight", "BottomLeft", "TopLeft")][string]$Position = "BottomRight",
    [ValidateSet("Small", "Medium", "Large")][string]$Size = "Small",
    [string]$FontFamily = "Segoe UI",
    [switch]$AlwaysOnTop = $false,
    [string]$AccentColorHex = "#0EA5E9",
    [string]$BgColorHex = "#1A202C",
    [string]$TextColorHex = "#F1F5F9",
    $ShowHost = $true,
    $ShowUser = $true,
    $ShowIP = $true,
    $ShowSerial = $true,
    [string]$BuildVersion = "1.0.0",
    [string]$LogPath = "C:\ProgramData\ITSupportOverlay\overlay.log"
)

# Convert boolean parameters flexibly
function Convert-ToBool ($val) {
    if ($val -is [bool]) { return $val }
    if ($val -eq 1 -or $val -eq "1" -or $val.ToString().ToLower() -eq "true") { return $true }
    return $false
}

$ShowHost   = Convert-ToBool $ShowHost
$ShowUser   = Convert-ToBool $ShowUser
$ShowIP     = Convert-ToBool $ShowIP
$ShowSerial = Convert-ToBool $ShowSerial

# --- Logging Helper ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")][string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $FormattedMessage = "[$TimeStamp] [$Level] [PID:$PID] $Message"
    
    switch ($Level) {
        "ERROR" { Write-Host $FormattedMessage -ForegroundColor Red }
        "WARN"  { Write-Host $FormattedMessage -ForegroundColor Yellow }
        default { Write-Host $FormattedMessage -ForegroundColor Cyan }
    }
    
    try {
        $TargetFile = $LogPath
        $LogDir = Split-Path -Path $TargetFile -Parent
        if (-not (Test-Path -Path $LogDir)) {
            $null = New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        }
        Add-Content -Path $TargetFile -Value $FormattedMessage -ErrorAction SilentlyContinue
    } catch {
        try {
            $Fallback = "$env:TEMP\ITSupportOverlay.log"
            Add-Content -Path $Fallback -Value $FormattedMessage -ErrorAction SilentlyContinue
        } catch {}
    }
}

Write-Log "--- Starting IT Support Desktop Overlay (Position: $Position, Desktop Mode: $([bool](-not $AlwaysOnTop))) ---" "INFO"

# --- In-Memory Safety Guard ---
if ([string]::IsNullOrEmpty($MyInvocation.MyCommand.Path) -or $MyInvocation.MyCommand.Path -match "^iex") {
    Write-Log "In-memory execution (via iex / WebString) detected. Aborting." "ERROR"
    Write-Error "In-memory execution is not supported. Please download and extract the package locally."
    exit 1
}

# --- Load Windows Forms & Drawing Assemblies ---
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
} catch {
    Write-Log "Failed to load WinForms/Drawing assemblies: $_" "ERROR"
    exit 1
}

# --- High-DPI Awareness ---
if (-not ([System.Management.Automation.PSTypeName]'DpiHelper').Type) {
    $DpiSignature = @"
using System;
using System.Runtime.InteropServices;

public class DpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
    try { Add-Type -TypeDefinition $DpiSignature -ErrorAction SilentlyContinue } catch {}
}
try { [DpiHelper]::SetProcessDPIAware() | Out-Null } catch {}

# --- Win32 Interop for Desktop Attachment & Extended Window Styles ---
if (-not ([System.Management.Automation.PSTypeName]'Win32OverlayHelper').Type) {
    $Win32Signature = @"
using System;
using System.Runtime.InteropServices;

public class Win32OverlayHelper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
"@
    try {
        Add-Type -TypeDefinition $Win32Signature -ErrorAction SilentlyContinue
    } catch {}
}

# --- System Information Helpers ---
function Get-PrimaryIPv4 {
    try {
        $ip = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "Wi-Fi*", "Ethernet*" -ErrorAction SilentlyContinue | 
               Where-Object { $_.IPAddress -notlike "169.254*" -and $_.IPAddress -ne "127.0.0.1" } | 
               Select-Object -First 1).IPAddress
        if ($ip) { return $ip }
    } catch {}
    return "127.0.0.1 (Sandbox/Local)"
}

function Get-SerialNumber {
    try {
        $sn = (Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber
        if ($sn -and $sn.Trim() -ne "") { return $sn.Trim() }
    } catch {}
    return "Sandbox-VM"
}

function Parse-Color ([string]$hex, [System.Drawing.Color]$defaultColor) {
    try {
        return [System.Drawing.ColorTranslator]::FromHtml($hex)
    } catch {
        return $defaultColor
    }
}

$AccentColor = Parse-Color $AccentColorHex ([System.Drawing.Color]::FromArgb(14, 165, 233))
$BgColor     = Parse-Color $BgColorHex ([System.Drawing.Color]::FromArgb(26, 32, 44))
$TextColor   = Parse-Color $TextColorHex ([System.Drawing.Color]::FromArgb(241, 245, 249))

$ComputerName = $env:COMPUTERNAME
$UserName     = $env:USERNAME
$IPAddress    = Get-PrimaryIPv4
$SerialNumber = Get-SerialNumber

# --- UI Setup ---
try {
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
    $form.ShowInTaskbar   = $false
    $form.TopMost         = [bool]$AlwaysOnTop
    $form.BackColor       = $BgColor
    $form.Opacity         = 0.92
    # Size Presets Calculation
    switch ($Size) {
        "Medium" {
            $formWidth    = 380
            $formHeight   = 200
            $fontSize     = 10.5
            $dashSeparator = "----------------------------------------------------"
        }
        "Large" {
            $formWidth    = 460
            $formHeight   = 240
            $fontSize     = 12.0
            $dashSeparator = "------------------------------------------------------------------"
        }
        default { # Small
            $formWidth    = 310
            $formHeight   = 160
            $fontSize     = 9.0
            $dashSeparator = "-----------------------------------------------"
        }
    }

    $form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)

    # Position Calculation
    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $padding = 20

    switch ($Position) {
        "TopRight" {
            $posX = $workingArea.Right - $form.Width - $padding
            $posY = $workingArea.Top + $padding
        }
        "BottomLeft" {
            $posX = $workingArea.Left + $padding
            $posY = $workingArea.Bottom - $form.Height - $padding
        }
        "TopLeft" {
            $posX = $workingArea.Left + $padding
            $posY = $workingArea.Top + $padding
        }
        default { # BottomRight
            $posX = $workingArea.Right - $form.Width - $padding
            $posY = $workingArea.Bottom - $form.Height - $padding
        }
    }

    $form.Location = New-Object System.Drawing.Point($posX, $posY)

    # Panel for Padding
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $panel.Padding   = New-Object System.Windows.Forms.Padding(12)
    $panel.BackColor = $BgColor
    $form.Controls.Add($panel)

    # Top Accent Bar
    $accent = New-Object System.Windows.Forms.Panel
    $accent.Height    = 4
    $accent.Dock      = [System.Windows.Forms.DockStyle]::Top
    $accent.BackColor = $AccentColor
    $panel.Controls.Add($accent)

    # Build Text Lines
    $lines = @()
    $lines += $Title.ToUpper()
    $lines += $dashSeparator

    if ($FieldOrder -eq "EmailFirst") {
        if (-not [string]::IsNullOrWhiteSpace($SupportEmail)) { $lines += "Email:  $SupportEmail" }
        if (-not [string]::IsNullOrWhiteSpace($SupportPhone)) { $lines += "Phone:  $SupportPhone" }
    } else {
        if (-not [string]::IsNullOrWhiteSpace($SupportPhone)) { $lines += "Phone:  $SupportPhone" }
        if (-not [string]::IsNullOrWhiteSpace($SupportEmail)) { $lines += "Email:  $SupportEmail" }
    }
    if (-not [string]::IsNullOrWhiteSpace($SupportHours)) {
        $lines += "Hours:  $SupportHours"
    }
    $lines += ""

    $sysFields = @()
    if ($ShowHost) { $sysFields += "Host: $ComputerName" }
    if ($ShowUser) { $sysFields += "User: $UserName" }
    if ($sysFields.Count -gt 0) { $lines += ($sysFields -join "  |  ") }

    $detailFields = @()
    if ($ShowIP)     { $detailFields += "IP: $IPAddress" }
    if ($ShowSerial) { $detailFields += "Serial: $SerialNumber" }
    if ($detailFields.Count -gt 0) { $lines += ($detailFields -join "  |  ") }

    # Content Text Label
    $label = New-Object System.Windows.Forms.Label
    $label.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $label.ForeColor = $TextColor
    $selectedFont    = try { New-Object System.Drawing.Font($FontFamily, $fontSize, [System.Drawing.FontStyle]::Regular) } catch { New-Object System.Drawing.Font("Segoe UI", $fontSize) }
    $label.Font      = $selectedFont
    $label.Padding   = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $label.Text      = ($lines -join "`n")
    $panel.Controls.Add($label)

    # --- Apply Desktop Background Z-Ordering & Extended Styles On Load ---
    $form.Add_Load({
        try {
            $GWL_EXSTYLE        = -20
            $WS_EX_TRANSPARENT  = 0x00000020
            $WS_EX_TOOLWINDOW   = 0x00000080
            $WS_EX_NOACTIVATE   = 0x08000000

            $hasWin32 = try { [bool]([System.Management.Automation.PSTypeName]'Win32OverlayHelper').Type } catch { $false }
            if ($hasWin32) {
                $currentStyle = [Win32OverlayHelper]::GetWindowLong($form.Handle, $GWL_EXSTYLE)
                $newStyle = ($currentStyle -bor $WS_EX_TRANSPARENT -bor $WS_EX_TOOLWINDOW -bor $WS_EX_NOACTIVATE)
                $null = [Win32OverlayHelper]::SetWindowLong($form.Handle, $GWL_EXSTYLE, $newStyle)

                if (-not $AlwaysOnTop) {
                    $Progman = [Win32OverlayHelper]::FindWindow("Progman", $null)
                    if ($Progman -ne [IntPtr]::Zero) {
                        $null = [Win32OverlayHelper]::SetParent($form.Handle, $Progman)
                    }

                    $HWND_BOTTOM      = [IntPtr]1
                    $SWP_NOSIZE       = 0x0001
                    $SWP_NOMOVE       = 0x0002
                    $SWP_NOACTIVATE   = 0x0010
                    $SWP_FRAMECHANGED = 0x0020
                    $flags = $SWP_NOSIZE -bor $SWP_NOMOVE -bor $SWP_NOACTIVATE -bor $SWP_FRAMECHANGED
                    $null = [Win32OverlayHelper]::SetWindowPos($form.Handle, $HWND_BOTTOM, 0, 0, 0, 0, $flags)
                }
            }

            Write-Log "Overlay window styles applied successfully." "INFO"
        } catch {
            Write-Log "Non-fatal error setting window Z-order in Form_Load: $_" "WARN"
        }
    })

    $form.Show()
    $form.Refresh()
    [System.Windows.Forms.Application]::Run($form)

} catch {
    Write-Log "Fatal error in overlay loop: $_" "ERROR"
    try {
        [System.Windows.Forms.MessageBox]::Show("Error starting desktop overlay:`n`n$_", "Overlay Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {}
    Start-Sleep -Seconds 5
    exit 1
}


