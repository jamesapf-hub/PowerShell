<#
.SYNOPSIS
    A sample script showing the usage of the repository's logging system.
.DESCRIPTION
    Demonstrates importing the Logging module and writing different levels of logs.
#>
param (
    [Parameter(Mandatory = $false)]
    [string]$TargetName = "World"
)

# 1. Locate and import the Logging module dynamically
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CandidatePaths = @(
    (Join-Path $ScriptDir "..\Modules\Logging\Logging.psm1"),  # Script in Scripts/ or Templates/
    (Join-Path $ScriptDir "Modules\Logging\Logging.psm1"),     # Script in Repo root
    (Join-Path $ScriptDir "..\..\Modules\Logging\Logging.psm1") # Script in a nested subdirectory
)

$ModuleImported = $false
foreach ($Path in $CandidatePaths) {
    if (Test-Path $Path) {
        Import-Module $Path -Force -ErrorAction SilentlyContinue
        $ModuleImported = $true
        break
    }
}

if (-not $ModuleImported) {
    Write-Error "CRITICAL: The Logging module (Logging.psm1) could not be located. Please make sure the Modules/Logging folder exists in the repository."
    exit 1
}

# 2. Initialize logging (creates C:\Log\<ProjectName> if not exists, determines file path)
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = $MyInvocation.MyCommand.Path
$LogFile = Initialize-ScriptLogging -ScriptPath $ScriptPath

# 3. Main script logic wrapped in try-catch for automatic logging
try {
    Write-ScriptLog -Message "Starting sample script execution." -LogFile $LogFile -Level INFO
    
    Write-ScriptLog -Message "Hello, $TargetName!" -LogFile $LogFile -Level INFO
    
    # Simulating a warning
    Write-ScriptLog -Message "This is a demonstration of a warning message." -LogFile $LogFile -Level WARNING

    # Simulating a successful step
    Write-ScriptLog -Message "Sample operations completed successfully." -LogFile $LogFile -Level SUCCESS
}
catch {
    $ErrorMessage = $_.Exception.Message
    Write-ScriptLog -Message "CRITICAL ERROR: $ErrorMessage" -LogFile $LogFile -Level ERROR
    exit 1
}
