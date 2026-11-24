<#
.SYNOPSIS
    Configures Windows system, user, welcome screen, and keyboard settings to English (United Kingdom) (en-GB).

.DESCRIPTION
    Installs the en-GB Windows language pack (core display pack only by default via -ExcludeFeatures for speed and reliability,
    with optional Features on Demand support via -IncludeFeatures or DATTO_INCLUDE_FEATURES), sets System Preferred UI Language,
    user culture, system locale, UK GeoID (242), user language list, and UI override, disables cloud language sync, and copies
    settings to the Welcome Screen and New User templates.

.PARAMETER IncludeFeatures
    Also installs optional language Features on Demand (BasicTyping, Handwriting, Speech, TextToSpeech, OCR).
    Omitted by default so Install-Language runs with -ExcludeFeatures and avoids slow or partial FoD failures.
    Can also be enabled by setting environment variable DATTO_INCLUDE_FEATURES=true (or 1 / yes).

.EXAMPLE
    .\Set-SystemLanguageUK.ps1
    Installs the core en-GB language pack (if missing) and applies all UK regional and keyboard settings.

.EXAMPLE
    .\Set-SystemLanguageUK.ps1 -IncludeFeatures
    Installs the en-GB language pack along with optional Features on Demand, falling back to core-only if FoDs fail.

.NOTES
    Script Name   : Set-SystemLanguageUK.ps1
    Author        : JP
    Version       : 1.1.0
    Created       : 29/07/26
    Last Modified : 24/11/26
    Log Path      : C:\Logs\Set-SystemLanguageUK\LanguageSetup.txt
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeFeatures
)

$installOptionalFeatures = $IncludeFeatures.IsPresent -or (
    -not [string]::IsNullOrWhiteSpace($env:DATTO_INCLUDE_FEATURES) -and
    $env:DATTO_INCLUDE_FEATURES.Trim() -match '^(1|true|yes)$'
)

# Standard UK format timestamped logger with %LOCALAPPDATA% fallback
$script:PrimaryLogFolder = Join-Path -Path $env:SystemDrive -ChildPath "Logs\Set-SystemLanguageUK"
$script:FallbackLogFolder = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Set-SystemLanguageUK"
} else {
    $null
}
$script:LogFolder = $script:PrimaryLogFolder
$script:LogFile = Join-Path -Path $script:LogFolder -ChildPath "LanguageSetup.txt"
$script:FallbackLogFile = if ($script:FallbackLogFolder) {
    Join-Path -Path $script:FallbackLogFolder -ChildPath "LanguageSetup.txt"
} else {
    $null
}

if (-not $WhatIfPreference) {
    if (-not (Test-Path -Path $script:LogFolder)) {
        try {
            New-Item -ItemType Directory -Path $script:LogFolder -Force -ErrorAction Stop | Out-Null
        } catch {
            if ($script:FallbackLogFolder) {
                try {
                    if (-not (Test-Path -Path $script:FallbackLogFolder)) {
                        New-Item -ItemType Directory -Path $script:FallbackLogFolder -Force -ErrorAction Stop | Out-Null
                    }
                    $script:LogFolder = $script:FallbackLogFolder
                    $script:LogFile = $script:FallbackLogFile
                } catch {
                    $script:LogFile = $null
                }
            } else {
                $script:LogFile = $null
            }
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [string]$Message
    )

    $timeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $entry = "[ $timeStamp ] $Message"

    if (-not $WhatIfPreference -and -not [string]::IsNullOrWhiteSpace($script:LogFile)) {
        try {
            $entry | Out-File -FilePath $script:LogFile -Append -Encoding UTF8 -ErrorAction Stop
        } catch {
            if ($script:FallbackLogFile -and $script:LogFile -ne $script:FallbackLogFile) {
                try {
                    if (-not (Test-Path -Path $script:FallbackLogFolder)) {
                        New-Item -ItemType Directory -Path $script:FallbackLogFolder -Force -ErrorAction Stop | Out-Null
                    }
                    $script:LogFolder = $script:FallbackLogFolder
                    $script:LogFile = $script:FallbackLogFile
                    $entry | Out-File -FilePath $script:LogFile -Append -Encoding UTF8 -ErrorAction Stop
                } catch {
                    $script:LogFile = $null
                }
            }
        }
    }

    Write-Output $entry
}

function Test-ShouldProcess {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Target,
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    # $PSCmdlet can be $null at script scope when executed via "irm ... | iex"
    if ($PSCmdlet) {
        return $PSCmdlet.ShouldProcess($Target, $Action)
    }

    return (-not $WhatIfPreference)
}

function ConvertTo-HexErrorCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $null
    }

    if ($Message -match 'ErrorCode:\s*(-?\d+)') {
        [long]$parsed = 0
        if ([long]::TryParse($Matches[1], [ref]$parsed)) {
            return ('0x{0:X8}' -f ([uint32]($parsed -band 0xFFFFFFFFL)))
        }
    }

    return $null
}

function Get-EnGbLanguageState {
    [CmdletBinding()]
    param()

    $getInstalledCmd = Get-Command -Name "Get-InstalledLanguage" -ErrorAction SilentlyContinue
    if (-not $getInstalledCmd) {
        return $null
    }

    try {
        $entry = Get-InstalledLanguage -ErrorAction Stop |
            Where-Object { $_.LanguageId -eq "en-GB" } |
            Select-Object -First 1
    } catch {
        return $null
    }

    if (-not $entry) {
        return [pscustomobject]@{
            Found            = $false
            HasLanguagePack  = $false
            LanguagePacks    = "None"
            LanguageFeatures = "None"
        }
    }

    [int]$packsInt = 0
    try {
        $packsInt = [int]$entry.LanguagePacks
    } catch {
        $packsInt = 0
    }

    return [pscustomobject]@{
        Found            = $true
        HasLanguagePack  = ($packsInt -ne 0)
        LanguagePacks    = [string]$entry.LanguagePacks
        LanguageFeatures = [string]$entry.LanguageFeatures
    }
}

function Write-LanguageInstallDiagnostic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$HexErrorCode
    )

    try {
        $auPolicy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue
        $servicingPolicy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Servicing" -ErrorAction SilentlyContinue
        if ($auPolicy -and $auPolicy.UseWUServer -eq 1 -and (-not $servicingPolicy -or $servicingPolicy.RepairContentServerSource -ne 2)) {
            Write-Log "DIAGNOSTIC: WSUS policy (UseWUServer=1) is active without 'Download repair content and optional features directly from Windows Update' (RepairContentServerSource=2), which commonly blocks language pack downloads."
        }
    } catch {
        Write-Verbose "WSUS policy check skipped: $($_.Exception.Message)"
    }

    try {
        $cbsReboot = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
        $wuReboot = Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
        if ($cbsReboot -or $wuReboot) {
            Write-Log "DIAGNOSTIC: A Windows pending reboot was detected (CBS/Windows Update). Reboot the workstation and re-run this script."
        }
    } catch {
        Write-Verbose "Pending reboot check skipped: $($_.Exception.Message)"
    }

    try {
        $editionId = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "EditionID" -ErrorAction SilentlyContinue).EditionID
        if ($editionId -match "SingleLanguage") {
            Write-Log "DIAGNOSTIC: Windows EditionID is '$editionId'. Single Language editions do not permit installing additional display language packs."
        }
    } catch {
        Write-Verbose "EditionID check skipped: $($_.Exception.Message)"
    }

    if ($HexErrorCode -like "0x800F*") {
        Write-Log "DIAGNOSTIC: $HexErrorCode is a Component-Based Servicing (CBS) error. Run 'DISM /Online /Cleanup-Image /RestoreHealth', reboot, and retry."
    }
}

$modeLabel = if ($installOptionalFeatures) {
    "Full language pack + optional Features on Demand"
} else {
    "Core language pack only (-ExcludeFeatures)"
}

Write-Log "=== System Language Configuration Started ==="
Write-Log "Target Language: en-GB (English UK)"
Write-Log "Install Mode: $modeLabel"

try {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Log "WARNING: Script is not running elevated as Administrator. System-wide changes and language pack installation may fail."
    }
} catch {
    Write-Verbose "Elevation check skipped: $($_.Exception.Message)"
}

$enGbPackInstalled = $false

# 1. Download and Install Language Pack (Windows 10 2004+ / Windows 11)
$installLangCmd = Get-Command -Name "Install-Language" -ErrorAction SilentlyContinue
if ($installLangCmd) {
    $preState = Get-EnGbLanguageState
    if ($preState -and $preState.HasLanguagePack -and -not $installOptionalFeatures) {
        $enGbPackInstalled = $true
        Write-Log "INFO: en-GB language pack is already installed (Packs: $($preState.LanguagePacks); Features: $($preState.LanguageFeatures)). Skipping download."
    } else {
        if ($preState -and $preState.HasLanguagePack) {
            $enGbPackInstalled = $true
        }

        if (Test-ShouldProcess -Target "en-GB" -Action "Install Windows language pack") {
            try {
                Write-Log "Attempting to install en-GB language pack ($modeLabel)..."
                $installParams = @{
                    Language    = "en-GB"
                    ErrorAction = "Stop"
                }
                if (-not $installOptionalFeatures) {
                    $installParams["ExcludeFeatures"] = $true
                }

                Install-Language @installParams | Out-Null
                $enGbPackInstalled = $true
                $postState = Get-EnGbLanguageState
                if ($postState) {
                    Write-Log "SUCCESS: Language pack en-GB installed (Packs: $($postState.LanguagePacks); Features: $($postState.LanguageFeatures))."
                } else {
                    Write-Log "SUCCESS: Language pack en-GB installed."
                }
            } catch {
                $errorText = [string]$_
                $hexCode = ConvertTo-HexErrorCode -Message $errorText
                $postState = Get-EnGbLanguageState

                if ($installOptionalFeatures -and (-not $postState -or -not $postState.HasLanguagePack)) {
                    Write-Log "WARNING: Full en-GB install failed ($errorText). Retrying core language pack with -ExcludeFeatures..."
                    try {
                        Install-Language -Language "en-GB" -ExcludeFeatures -ErrorAction Stop | Out-Null
                        $postState = Get-EnGbLanguageState
                    } catch {
                        $retryErrorText = [string]$_
                        $retryHex = ConvertTo-HexErrorCode -Message $retryErrorText
                        if ($retryHex) {
                            $hexCode = $retryHex
                        }
                        $errorText = $retryErrorText
                        $postState = Get-EnGbLanguageState
                    }
                }

                if ($postState -and $postState.HasLanguagePack) {
                    $enGbPackInstalled = $true
                    $hexSummary = if ($hexCode) { " ($hexCode)" } else { "" }
                    Write-Log "WARNING: Core en-GB language pack is installed (Packs: $($postState.LanguagePacks); Features: $($postState.LanguageFeatures)), though optional features reported a partial/feature error$hexSummary. Continuing configuration."
                } else {
                    $hexSummary = if ($hexCode) { " [Hex: $hexCode]" } else { "" }
                    Write-Log "ERROR: Failed to install en-GB language pack$hexSummary. Reason: $errorText"
                    Write-LanguageInstallDiagnostic -HexErrorCode $hexCode
                }
            }
        }
    }
} else {
    Write-Log "INFO: Install-Language cmdlet not available. Skipping language pack download."
}

# 2. Set System Preferred UI Language (when core en-GB language pack is confirmed)
$setPreferredCmd = Get-Command -Name "Set-SystemPreferredUILanguage" -ErrorAction SilentlyContinue
if ($setPreferredCmd) {
    if ($enGbPackInstalled) {
        try {
            $currentPreferred = $null
            $getPreferredCmd = Get-Command -Name "Get-SystemPreferredUILanguage" -ErrorAction SilentlyContinue
            if ($getPreferredCmd) {
                $currentPreferred = Get-SystemPreferredUILanguage -ErrorAction SilentlyContinue
            }

            if ($currentPreferred -eq "en-GB") {
                Write-Log "INFO: System Preferred UI Language is already en-GB."
            } elseif (Test-ShouldProcess -Target "en-GB" -Action "Set System Preferred UI Language") {
                Set-SystemPreferredUILanguage -Language "en-GB" -ErrorAction Stop
                Write-Log "SUCCESS: System Preferred UI Language set to en-GB."
            }
        } catch {
            Write-Log "ERROR: Failed to set System Preferred UI Language. Reason: $_"
        }
    } else {
        Write-Log "INFO: Skipping Set-SystemPreferredUILanguage because en-GB language pack is not confirmed installed."
    }
}

# 3. Set User Culture (Current Context)
try {
    if (Test-ShouldProcess -Target "en-GB" -Action "Set User Culture") {
        Set-Culture -CultureInfo "en-GB" -ErrorAction Stop
        Write-Log "SUCCESS: Culture set to en-GB."
    }
} catch {
    Write-Log "ERROR: Failed to set culture. Reason: $_"
}

# 4. Set System Locale (Non-Unicode programs)
try {
    if (Test-ShouldProcess -Target "en-GB" -Action "Set System Locale") {
        Set-WinSystemLocale -SystemLocale "en-GB" -ErrorAction Stop
        Write-Log "SUCCESS: System Locale set to en-GB."
    }
} catch {
    Write-Log "ERROR: Failed to set System Locale. Reason: $_"
}

# 5. Set Home Location (GeoID 242 is United Kingdom)
try {
    if (Test-ShouldProcess -Target "GeoID 242 (United Kingdom)" -Action "Set Home Location") {
        Set-WinHomeLocation -GeoId 242 -ErrorAction Stop
        Write-Log "SUCCESS: Home Location GeoID set to 242 (UK)."
    }
} catch {
    Write-Log "ERROR: Failed to set Home Location. Reason: $_"
}

# 6. Set Language List and UI Override
try {
    if (Test-ShouldProcess -Target "en-GB" -Action "Set User Language List and UI Language Override") {
        $languageList = New-WinUserLanguageList -Language "en-GB"
        Set-WinUserLanguageList -LanguageList $languageList -Force -ErrorAction Stop
        try {
            Set-WinUILanguageOverride -Language "en-GB" -ErrorAction Stop
        } catch {
            Write-Log "INFO: UI language override warning: $_"
        }
        Write-Log "SUCCESS: User language list and UI override set to en-GB."
    }
} catch {
    Write-Log "ERROR: Failed to set user language list. Reason: $_"
}

# 7. Disable Language Synchronization (Prevents Entra ID / MSA from reverting settings)
try {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\SettingSync\Groups\Language"
    if (Test-ShouldProcess -Target $regPath -Action "Disable Cloud Language Synchronization") {
        if (-not (Test-Path -Path $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "Enabled" -Value 0 -Type DWord -Force -ErrorAction Stop
        Write-Log "SUCCESS: Cloud language synchronization disabled in Registry."
    }
} catch {
    Write-Log "ERROR: Failed to disable language sync. Reason: $_"
}

# 8. Copy Settings to Welcome Screen (Logon page) and New User Templates
$copyIntlCmd = Get-Command -Name "Copy-UserInternationalSettingsToSystem" -ErrorAction SilentlyContinue
if ($copyIntlCmd) {
    try {
        if (Test-ShouldProcess -Target "Welcome Screen and New User Templates" -Action "Copy International Settings") {
            Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true -ErrorAction Stop
            Write-Log "SUCCESS: International settings copied to Welcome Screen (logon page) and New User Templates."
        }
    } catch {
        Write-Log "ERROR: Failed to copy international settings to system accounts. Reason: $_"
    }
} else {
    Write-Log "INFO: Copy-UserInternationalSettingsToSystem cmdlet not available on this OS version."
}

Write-Log "=== System Language Configuration Finished ==="
Write-Log "NOTE: A system reboot is required for language changes to take full effect on the logon screen."
