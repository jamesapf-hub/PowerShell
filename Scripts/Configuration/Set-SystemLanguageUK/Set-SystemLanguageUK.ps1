# Standard UK format timestamped logger
$LogFolder = "$env:SystemDrive\Logs\Set-SystemLanguageUK"
$LogFile   = "$LogFolder\LanguageSetup.txt"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

function Write-Log ($Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    "[ $TimeStamp ] $Message" | Out-File -FilePath $LogFile -Append
    Write-Output "[ $TimeStamp ] $Message"
}

Write-Log "=== System Language Configuration Started ==="
Write-Log "Target Language: en-GB (English UK)"

try {
    # 1. Download and Install Language Pack (Windows 10 2004+ / Windows 11)
    if (Get-Command Install-Language -ErrorAction SilentlyContinue) {
        Write-Log "Attempting to install en-GB language pack..."
        Install-Language -Language "en-GB" -ErrorAction Stop
        Write-Log "SUCCESS: Language pack en-GB installed."
    } else {
        Write-Log "INFO: Install-Language cmdlet not available. Skipping language pack download."
    }
} catch {
    Write-Log "ERROR: Failed to install language pack. Reason: $_"
}

try {
    # 2. Set User Culture (Current Context)
    Set-Culture -CultureInfo "en-GB" -ErrorAction Stop
    Write-Log "SUCCESS: Culture set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set culture. Reason: $_"
}

try {
    # 3. Set System Locale (Non-Unicode programs)
    Set-WinSystemLocale -SystemLocale "en-GB" -ErrorAction Stop
    Write-Log "SUCCESS: System Locale set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set System Locale. Reason: $_"
}

try {
    # 4. Set Home Location (GeoID 242 is United Kingdom)
    Set-WinHomeLocation -GeoId 242 -ErrorAction Stop
    Write-Log "SUCCESS: Home Location GeoID set to 242 (UK)."
} catch {
    Write-Log "ERROR: Failed to set Home Location. Reason: $_"
}

try {
    # 5. Set Language List and UI Override
    $LanguageList = New-WinUserLanguageList -Language "en-GB"
    Set-WinUserLanguageList -LanguageList $LanguageList -Force -ErrorAction Stop
    Set-WinUILanguageOverride -Language "en-GB" -ErrorAction SilentlyContinue
    Write-Log "SUCCESS: User language list and UI override set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set user language list. Reason: $_"
}

try {
    # 6. Disable Language Synchronization (Prevents Entra ID / MSA from reverting settings)
    $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\SettingSync\Groups\Language"
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $RegPath -Name "Enabled" -Value 0 -Force -ErrorAction Stop
    Write-Log "SUCCESS: Cloud language synchronization disabled in Registry."
} catch {
    Write-Log "ERROR: Failed to disable language sync. Reason: $_"
}

try {
    # 7. Copy Settings to Welcome Screen (Logon page) and New User Templates
    # This replicates the settings to HKEY_USERS\.DEFAULT (Logon screen) and the default profile template.
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUserTemplates $true -ErrorAction Stop
    Write-Log "SUCCESS: International settings copied to Welcome Screen (logon page) and New User Templates."
} catch {
    Write-Log "ERROR: Failed to copy international settings to system accounts. Reason: $_"
}

Write-Log "=== System Language Configuration Finished ==="
Write-Log "NOTE: A system reboot is required for language changes to take full effect on the logon screen."
exit 0
