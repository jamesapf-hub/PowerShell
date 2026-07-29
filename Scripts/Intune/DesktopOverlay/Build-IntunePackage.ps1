<#
.SYNOPSIS
    Builds the .intunewim Win32 Package for Microsoft Intune Deployment
.DESCRIPTION
    Downloads IntuneWinAppUtil.exe from Microsoft if missing, packages the src folder,
    and produces the output .intunewim file ready for Intune portal upload.
#>

$ErrorActionPreference = "Stop"

# --- In-Memory Guard for Bundled Packages ---
if ([string]::IsNullOrEmpty($PSScriptRoot) -or $PSScriptRoot -match "^iex") {
    Write-Host "ERROR: In-memory execution (via iex / irm) is not supported for bundled packages." -ForegroundColor Red
    Write-Error "Build-IntunePackage.ps1 relies on local relative files (src/ directory and IntuneWinAppUtil.exe). Please download and extract the repository package locally before executing."
    exit 1
}

$WorkspaceRoot = $PSScriptRoot
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
} else {
    Write-Error "IntuneWinAppUtil failed with exit code $($Process.ExitCode):`n$Output`n$ErrorOutput"
}
