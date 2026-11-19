<#
.SYNOPSIS
    Scaffolds a new wrapped-folder PowerShell tool matching repository standards.
.DESCRIPTION
    Creates Scripts/<Category>/<ToolName>/ with <ToolName>.ps1, README.md, .gitignore,
    and an optional Start-Gui.bat launcher.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Azure', 'Clean Up Tools', 'Configuration', 'Datto', 'Intune', 'Kiosk', 'Microsoft 365', 'Networking', 'Sample', 'Software')]
    [string]$Category,

    [Parameter(Mandatory = $true)]
    [string]$ToolName,

    [Parameter(Mandatory = $false)]
    [string]$Synopsis = "Automated PowerShell utility for $ToolName.",

    [Parameter(Mandatory = $false)]
    [switch]$Gui,

    [Parameter(Mandatory = $false)]
    [switch]$Intune
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..\..\..")
$targetDir = Join-Path $repoRoot "Scripts\$Category\$ToolName"

if (-not (Test-Path -Path $targetDir)) {
    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
}

$scriptPath = Join-Path $targetDir "$ToolName.ps1"
$readmePath = Join-Path $targetDir "README.md"
$gitignorePath = Join-Path $targetDir ".gitignore"
$dateStr = Get-Date -Format "dd/MM/yyyy"

$staBlock = if ($Gui) {
@"

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    `$relaunchPath = if (`$PSScriptRoot) { Join-Path `$PSScriptRoot `$MyInvocation.MyCommand.Name } else { `$PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"`$relaunchPath`""
    exit
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
"@
} else { "" }

if (-not (Test-Path -Path $scriptPath)) {
    $psContent = @"
<#
.SYNOPSIS
    $Synopsis

.DESCRIPTION
    $Synopsis
    Writes diagnostic execution logs to C:\Logs\$ToolName (with fallback to %LOCALAPPDATA%\Logs\$ToolName).

.EXAMPLE
    .\$ToolName.ps1

.NOTES
    Script Name   : $ToolName.ps1
    Author        : JP
    Date          : $dateStr
    Version       : 1.0.0
    Log Path      : C:\Logs\$ToolName\$ToolName.log
#>
[CmdletBinding(SupportsShouldProcess = `$true)]
param()
$staBlock
`$ErrorActionPreference = 'Stop'

function Write-ToolLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = `$true)][string]`$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]`$Level = 'INFO'
    )
    `$logDir = "C:\Logs\$ToolName"
    if (-not (Test-Path -Path `$logDir)) {
        try {
            New-Item -Path `$logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            `$logDir = Join-Path `$env:LOCALAPPDATA "Logs\$ToolName"
            if (-not (Test-Path -Path `$logDir)) {
                New-Item -Path `$logDir -ItemType Directory -Force | Out-Null
            }
        }
    }
    `$timestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    `$entry = "[`$timestamp] [`$Level] `$Message"
    Add-Content -Path (Join-Path `$logDir "$ToolName.log") -Value `$entry -Encoding UTF8
    Write-Verbose `$entry
}

try {
    Write-ToolLog -Message "Starting $ToolName..." -Level INFO
    # TODO: Implement core tool logic here
    Write-ToolLog -Message "$ToolName completed successfully." -Level SUCCESS
} catch {
    Write-ToolLog -Message "Unhandled error: `$(`$_.Exception.Message)" -Level ERROR
    throw
}
"@
    Set-Content -Path $scriptPath -Value $psContent -Encoding UTF8
}

if ($Gui) {
    $batPath = Join-Path $targetDir "Start-Gui.bat"
    if (-not (Test-Path -Path $batPath)) {
        $batContent = @"
@echo off
title $ToolName
echo ============================================================
echo Launching $ToolName GUI...
echo ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0$ToolName.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script closed with exit code %ERRORLEVEL%. Press any key to exit.
    pause
)
"@
        Set-Content -Path $batPath -Value $batContent -Encoding ASCII
    }
}

if (-not (Test-Path -Path $gitignorePath)) {
    Set-Content -Path $gitignorePath -Value "Output/`r`n*.intunewin`r`n" -Encoding UTF8
}

if (-not (Test-Path -Path $readmePath)) {
    $fastExecBlock = if ($Gui -or $Intune) {
@"
> [!WARNING]
> **Bundled Package Notice:**
> This tool is a multi-file package and cannot be executed in-memory via ``irm | iex``. Please download or clone the repository locally before executing.
>
> [!TIP]
> **Run locally in PowerShell (as Administrator):**
> Double-click ``Start-Gui.bat`` or run:
> ````powershell
> powershell.exe -ExecutionPolicy Bypass -File .\$ToolName.ps1
> ````
"@
    } else {
@"
> [!TIP]
> **Full Transparent GitHub Command:**
> ````powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/$Category/$ToolName/$ToolName.ps1")
> ````
"@
    }

    $readmeContent = @"
# $ToolName Guide

$Synopsis

## Overview

$Synopsis

### Key Features

- Standardized UK-timestamped diagnostic logging to ``C:\Logs\$ToolName\``.
- Structured error handling and PSScriptAnalyzer-compliant syntax.

## Prerequisites

- **OS Support:** Windows 10 (1909+) / Windows 11
- **PowerShell:** Windows PowerShell 5.1 or PowerShell 7+
- **Permissions:** Local Administrator privileges recommended

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Open an elevated PowerShell prompt or launch via ``Start-Gui.bat``.
2. Execute ``.\$ToolName.ps1``.
3. Review the execution output and log file in ``C:\Logs\$ToolName\$ToolName.log``.

### Logging & Diagnostics

| Log Path | Purpose |
| :--- | :--- |
| ``C:\Logs\$ToolName\$ToolName.log`` | Primary execution and error audit log |

## Fast Execute

$fastExecBlock
"@
    Set-Content -Path $readmePath -Value $readmeContent -Encoding UTF8
}

Write-Output "Scaffolded wrapped PowerShell tool at: $targetDir"
