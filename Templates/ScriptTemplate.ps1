<#
.SYNOPSIS
    Starter template for PowerShell scripts with built-in logging and error handling.
.DESCRIPTION
    This script automatically imports the local Logging module, creates a log file
    in C:\Log (or system drive log folder), and logs runtime progress and exceptions.
.PARAMETER SampleParam
    A sample input parameter for the script.
.EXAMPLE
    .\ScriptTemplate.ps1 -SampleParam "Hello"
#>
param (
    [Parameter(Mandatory = $false)]
    [string]$SampleParam = "DefaultValue"
)

# 1. Locate and import the Logging module dynamically
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CandidatePaths = @(
    (Join-Path $ScriptDir "..\Modules\Logging\Logging.psm1"),  # Script in Scripts/ or Templates/
    (Join-Path $ScriptDir "Modules\Logging\Logging.psm1"),     # Script in Repo root
    (Join-Path $ScriptDir "..\..\Modules\Logging\Logging.psm1"), # Script in a 2-level nested subdirectory
    (Join-Path $ScriptDir "..\..\..\Modules\Logging\Logging.psm1") # Script in a 3-level nested subdirectory
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
    Write-ScriptLog -Message "Starting script execution: $ScriptName" -LogFile $LogFile -Level INFO
    Write-ScriptLog -Message "Parameter SampleParam is set to: $SampleParam" -LogFile $LogFile -Level INFO

    # --- YOUR CODE GOES HERE ---
    Write-ScriptLog -Message "Performing script actions..." -LogFile $LogFile -Level INFO
    
    # Simulating successful action
    Write-ScriptLog -Message "Successfully completed the core tasks." -LogFile $LogFile -Level SUCCESS
    # ---------------------------

    Write-ScriptLog -Message "Script finished successfully." -LogFile $LogFile -Level INFO
}
catch {
    # Automatically log any unhandled runtime exceptions
    $ErrorMessage = $_.Exception.Message
    Write-ScriptLog -Message "CRITICAL ERROR: $ErrorMessage" -LogFile $LogFile -Level ERROR
    Write-ScriptLog -Message "Stack Trace: $($_.ScriptStackTrace)" -LogFile $LogFile -Level ERROR
    
    # Re-throw or exit depending on your environment needs
    exit 1
}
