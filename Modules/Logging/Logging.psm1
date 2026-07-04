# PowerShell Shared Logging Module
# Implements C:\Log folder creation and consistent timestamped logging

function Initialize-ScriptLogging {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $false)]
        [string]$LogDirectory
    )

    process {
        # Determine the script name and parent directory
        $ScriptName = Split-Path -Leaf $ScriptPath
        $ScriptDir = Split-Path -Parent $ScriptPath

        # 1. Discover Project Name (Git repo name, or fallback to parent folder name)
        $ProjectName = "Default"
        if ($ScriptDir -and (Test-Path $ScriptDir)) {
            $TempDir = $ScriptDir
            while ($TempDir) {
                if (Test-Path (Join-Path $TempDir ".git")) {
                    $ProjectName = Split-Path $TempDir -Leaf
                    break
                }
                $Parent = Split-Path $TempDir -Parent
                if ($Parent -eq $TempDir) { break }
                $TempDir = $Parent
            }
            if (-not $ProjectName) {
                # Fallback: parent folder, or parent of parent if named Scripts/Templates/Modules
                $FolderName = Split-Path $ScriptDir -Leaf
                if ($FolderName -eq "Scripts" -or $FolderName -eq "Templates" -or $FolderName -eq "Modules") {
                    $Parent = Split-Path $ScriptDir -Parent
                    if ($Parent) {
                        $ProjectName = Split-Path $Parent -Leaf
                    }
                } else {
                    $ProjectName = $FolderName
                }
            }
        }

        # 2. Determine Log Directory (C:\Log\<ProjectName>)
        if ([string]::IsNullOrEmpty($LogDirectory)) {
            $SystemDrive = $env:SystemDrive
            if ([string]::IsNullOrEmpty($SystemDrive)) {
                $SystemDrive = "C:"
            }
            $LogDirectory = Join-Path $SystemDrive "Log"
            $LogDirectory = Join-Path $LogDirectory $ProjectName
        }

        # Create C:\Log\<ProjectName> directory if it doesn't exist
        if (-not (Test-Path -Path $LogDirectory -PathType Container)) {
            try {
                $null = New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction Stop
                Write-Host "[INIT] Created log folder: $LogDirectory" -ForegroundColor Cyan
            }
            catch {
                Write-Error "Failed to create log folder at $LogDirectory. Error: $_"
                throw $_
            }
        }

        # Generate a dated log file path: C:\Log\<ProjectName>\<ScriptName>_dd-MM-yy.log
        $DateString = Get-Date -Format "dd-MM-yy"
        $SanitizedName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptName)
        $LogFileName = "${SanitizedName}_${DateString}.log"
        $LogFilePath = Join-Path $LogDirectory $LogFileName

        return $LogFilePath
    }
}

function Write-ScriptLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [string]$LogFile,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    process {
        $Timestamp = Get-Date -Format "dd\/MM\/yy HH:mm:ss"
        $LogEntry = "[$Timestamp] [$Level] $Message"

        # Write to console host with color formatting
        switch ($Level) {
            'ERROR'   { Write-Host $LogEntry -ForegroundColor Red }
            'WARNING' { Write-Host $LogEntry -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $LogEntry -ForegroundColor Green }
            Default   { Write-Host $LogEntry -ForegroundColor White }
        }

        # Write to log file
        try {
            Add-Content -Path $LogFile -Value $LogEntry -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not write to log file $LogFile. Error: $_"
        }
    }
}

Export-ModuleMember -Function Initialize-ScriptLogging, Write-ScriptLog
