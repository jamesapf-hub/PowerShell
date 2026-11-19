<#
.SYNOPSIS
    Builds the .intunewim Win32 Package for Microsoft Intune Deployment
.DESCRIPTION
    Downloads IntuneWinAppUtil.exe from Microsoft if missing, packages the src folder,
    and produces the output .intunewim file ready for Intune portal upload.
#>

$ErrorActionPreference = "Stop"

# --- In-Memory Guard for Bundled Packages ---
if ([string]::IsNullOrEmpty($MyInvocation.MyCommand.Path) -or $MyInvocation.MyCommand.Path -match "^iex") {
    throw "In-memory execution (via iex / irm) is not supported for bundled packages. Build-IntunePackage.ps1 relies on local relative files (src/ directory and IntuneWinAppUtil.exe). Please download and extract the repository package locally before executing."
}

$WorkspaceRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($WorkspaceRoot)) {
    $WorkspaceRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
$SrcFolder     = Join-Path -Path $WorkspaceRoot -ChildPath "src"
$OutputDir     = Join-Path -Path $WorkspaceRoot -ChildPath "output"
$ToolPath      = Join-Path -Path $WorkspaceRoot -ChildPath "IntuneWinAppUtil.exe"

Write-Host "=== IT Support Overlay Intune Package Builder ===" -ForegroundColor Cyan

# 1. Download IntuneWinAppUtil.exe if needed
if (-not (Test-Path -Path $ToolPath)) {
    Write-Host "IntuneWinAppUtil.exe not found locally. Downloading from Microsoft GitHub..." -ForegroundColor Yellow
    $DownloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $ToolPath -UseBasicParsing
        $sig = Get-AuthenticodeSignature -FilePath $ToolPath -ErrorAction Stop
        if ($sig.Status -ne 'Valid') {
            try { Remove-Item $ToolPath -Force -ErrorAction Stop } catch {}
            throw "Authenticode signature check failed for IntuneWinAppUtil.exe ($($sig.Status))"
        }
        Write-Host "Successfully downloaded IntuneWinAppUtil.exe" -ForegroundColor Green
    } catch {
        Write-Error "Failed to download IntuneWinAppUtil.exe: $_"
        exit 1
    }
} else {
    Write-Host "Found IntuneWinAppUtil.exe at $ToolPath" -ForegroundColor Green
}

# 2. Ensure Output Directory exists
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# 3. Execute IntuneWinAppUtil.exe
Write-Host "Packaging source directory ($SrcFolder) into .intunewim file..." -ForegroundColor Cyan

$SetupFile = "Install-ITOverlay.ps1"
$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
$ProcessInfo.FileName = $ToolPath
$ProcessInfo.Arguments = "-c `"$SrcFolder`" -s `"$SetupFile`" -o `"$OutputDir`" -q"
$ProcessInfo.UseShellExecute = $false
$ProcessInfo.RedirectStandardOutput = $true
$ProcessInfo.RedirectStandardError = $true

$Process = [System.Diagnostics.Process]::Start($ProcessInfo)
$Output = $Process.StandardOutput.ReadToEnd()
$ErrorOutput = $Process.StandardError.ReadToEnd()
$Process.WaitForExit()

if ($Process.ExitCode -eq 0) {
    Write-Host "Package creation successful!" -ForegroundColor Green
    $IntuneWimFile = Get-ChildItem -Path $OutputDir -Filter "*.intunewim" | Select-Object -First 1
    if ($IntuneWimFile) {
        Write-Host "Generated Intune Package: $($IntuneWimFile.FullName) ($([math]::Round($IntuneWimFile.Length / 1KB, 2)) KB)" -ForegroundColor Green
    }

    # Copy Custom Intune Detection Script to Output Directory
    $DetectSrc = Join-Path -Path $SrcFolder -ChildPath "Detect-ITOverlay.ps1"
    $DetectDst = Join-Path -Path $OutputDir -ChildPath "Detect-ITOverlay.ps1"
    if (Test-Path -Path $DetectSrc) {
        Copy-Item -Path $DetectSrc -Destination $DetectDst -Force
        Write-Host "Copied Intune Detection Script to Output: $DetectDst" -ForegroundColor Green
    }

    # Extract current overlay parameters for the Intune Description block
    $OverlayScriptPath = Join-Path -Path $SrcFolder -ChildPath "Start-ITOverlay.ps1"
    $OverlayTitle   = "IT SUPPORT HELPDESK"
    $OverlayPhone   = "0800 123 4567"
    $OverlayEmail   = "support@company.com"
    $OverlayVersion = "1.0.0"
    if (Test-Path -Path $OverlayScriptPath) {
        $ovRaw = try { Get-Content -Path $OverlayScriptPath -Raw -ErrorAction Stop } catch { $null }
        if ($ovRaw -match '\[string\]\$Title\s*=\s*"([^"]*)"')        { $OverlayTitle   = $Matches[1] }
        if ($ovRaw -match '\[string\]\$SupportPhone\s*=\s*"([^"]*)"') { $OverlayPhone   = $Matches[1] }
        if ($ovRaw -match '\[string\]\$SupportEmail\s*=\s*"([^"]*)"') { $OverlayEmail   = $Matches[1] }
        if ($ovRaw -match '\[string\]\$BuildVersion\s*=\s*"([^"]*)"') { $OverlayVersion = $Matches[1] }
    }

    $PackageFileName    = if ($IntuneWimFile) { $IntuneWimFile.Name } else { "Install-ITOverlay.intunewin" }
    $GeneratedTimestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $CreatedDateStr     = Get-Date -Format "dd/MM/yy"

    # Generate Intune-Deployment-Instructions.txt in Output Directory
    $InstructionsPath = Join-Path -Path $OutputDir -ChildPath "Intune-Deployment-Instructions.txt"
    $InstructionsContent = @"
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : $GeneratedTimestamp
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
$PackageFileName


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
IT Support Desktop Overlay ($OverlayVersion)

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
Deploys the non-interactive IT Support Desktop Overlay displaying
helpdesk contact details and live endpoint diagnostics on the desktop.
All runtime logs are saved to C:\Logs\DesktopOverlay\

PACKAGE DETAILS:
- Overlay Title : $OverlayTitle
- Support Phone : $OverlayPhone
- Support Email : $OverlayEmail
- Author        : JP
- App Version   : $OverlayVersion
- Created       : $CreatedDateStr

WORKSTATION LOG LOCATION:
C:\Logs\DesktopOverlay\

LOG FILES IN THIS FOLDER:
- Install Log   : C:\Logs\DesktopOverlay\Install.log
- Overlay Log   : C:\Logs\DesktopOverlay\Overlay.log
- Detection Log : C:\Logs\DesktopOverlay\Detection.log
- Uninstall Log : C:\Logs\DesktopOverlay\Uninstall.log
======================================================================

Publisher:
Pheneox

App Version:
$OverlayVersion


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install-ITOverlay.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall-ITOverlay.ps1

Install behavior:
System

Device restart behavior:
No specific action


STEP 4: REQUIREMENTS
------------------------------------------------------------------------
Operating system architecture:
64-bit

Minimum operating system:
Windows 10 1909


STEP 5: DETECTION RULES
------------------------------------------------------------------------
Rules format:
Use a custom detection script

Script file:
Detect-ITOverlay.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No


========================================================================
FILES IN THIS OUTPUT FOLDER
========================================================================
- $PackageFileName               : Upload to Intune Win32 App Package
- Detect-ITOverlay.ps1               : Upload to Intune Custom Detection Script
- Intune-Deployment-Instructions.txt : This deployment reference guide
========================================================================
"@
    Set-Content -Path $InstructionsPath -Value $InstructionsContent -Encoding UTF8
    Write-Host "Generated Deployment Instructions: $InstructionsPath" -ForegroundColor Green
} else {
    Write-Error "IntuneWinAppUtil failed with exit code $($Process.ExitCode):`n$Output`n$ErrorOutput"
}
