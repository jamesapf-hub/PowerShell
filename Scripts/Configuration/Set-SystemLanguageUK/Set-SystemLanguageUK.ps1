# Standard UK format timestamped logger
$LogFolder = "C:\Seriun\log"
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
    # 1. Set User Culture (Current Context)
    Set-Culture -CultureInfo "en-GB" -ErrorAction Stop
    Write-Log "SUCCESS: Culture set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set culture. Reason: $_"
}

try {
    # 2. Set System Locale (Non-Unicode programs)
    Set-WinSystemLocale -SystemLocale "en-GB" -ErrorAction Stop
    Write-Log "SUCCESS: System Locale set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set System Locale. Reason: $_"
}

try {
    # 3. Set Home Location (GeoID 242 is United Kingdom)
    Set-WinHomeLocation -GeoId 242 -ErrorAction Stop
    Write-Log "SUCCESS: Home Location GeoID set to 242 (UK)."
} catch {
    Write-Log "ERROR: Failed to set Home Location. Reason: $_"
}

try {
    # 4. Set Language List
    $LanguageList = New-WinUserLanguageList -Language "en-GB"
    Set-WinUserLanguageList -LanguageList $LanguageList -Force -ErrorAction Stop
    Write-Log "SUCCESS: User language list set to en-GB."
} catch {
    Write-Log "ERROR: Failed to set user language list. Reason: $_"
}

try {
    # 5. Copy Settings to Welcome Screen (Logon page) and New User Templates
    # This replicates the settings to HKEY_USERS\.DEFAULT (Logon screen) and the default profile template.
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUserTemplates $true -ErrorAction Stop
    Write-Log "SUCCESS: International settings copied to Welcome Screen (logon page) and New User Templates."
} catch {
    Write-Log "ERROR: Failed to copy international settings to system accounts. Reason: $_"
}

Write-Log "=== System Language Configuration Finished ==="
Write-Log "NOTE: A system reboot is required for language changes to take full effect on the logon screen."
exit 0
