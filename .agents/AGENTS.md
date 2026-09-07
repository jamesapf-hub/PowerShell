# PowerShell Repository Agent Guidelines

These rules and standards apply **STRICTLY AND EXCLUSIVELY** to the `PowerShell` repository (`jamesapf-hub/PowerShell`).

---

## 1. Context & Repository Isolation

1. **PowerShell Script Standards Isolation**: All guidelines regarding script structures, `README.md` generation, Fast Execute snippets (`irm https://phnx.it/... | iex`), and wrapped folder patterns belong exclusively to this repository.
2. **Zero Cross-Pollination**: Never introduce web project conventions, Cloudflare worker scripts, React components, or non-PowerShell requirements into this repository.
3. **Deployment Safety**: **ALWAYS** ask for explicit, fresh confirmation from the user before pushing commits to GitHub remote.

---

## 2. Directory Architecture (Wrapped Folder Pattern)

Every standalone script or application in this repository MUST reside inside its own dedicated wrapped directory:

```text
Scripts/
└── <Category>/
    └── <ToolName>/
        ├── <ToolName>.ps1           # Main PowerShell script or GUI application
        ├── README.md                # Comprehensive walkthrough and usage guide
        ├── Start-Gui.bat            # (Optional/Recommended) 1-click batch launcher for GUI tools
        ├── .gitignore               # (Optional) Ignores build outputs like Output/ or *.intunewin
        └── tools/                   # (Optional) Bundled dependencies (e.g. IntuneWinAppUtil.exe)
```

### Script Categories
Scripts must be filed under the appropriate category directory in `Scripts/`:
*   **Azure:** Azure VM metadata extraction, IMDS queries, cloud resource management, and automation.
*   **Clean Up Tools:** Utilities to clean system caches, remove temporary files, and reclaim disk space.
*   **Configuration:** OS parameter customization, registry settings, and profile configurations.
*   **Datto:** Datto RMM component scripts, policy triggers, and status checks.
*   **Intune:** Intune app packaging helpers, custom proactive remediations, and MDM policy templates.
*   **Kiosk:** Dedicated locking, runtime repair, and provisioning scripts for single-app kiosks.
*   **Microsoft 365:** User auditing, licensing checks, Teams rooms, and tenant management scripts.
*   **Networking:** Ping diagnostics, subnet scanners, DNS resolvers, and network status checks.
*   **Sample:** Reference templates demonstrating repository standards, logging, and error handling.
*   **Software:** Silent software deployment wrappers, updates, and application uninstallation scripts.

---

## 3. Script Authoring Standards

1. **Comment-Based Help Block**: Every script must begin with a complete comment-based help block containing:
   - `.SYNOPSIS`: Clear one-line summary of the script.
   - `.DESCRIPTION`: Detailed description of the functionality, parameters, and workflow.
   - `.PARAMETER`: Documentation for each parameter.
   - `.EXAMPLE`: Working command-line examples.
   - `.NOTES`: Author, date, version, and logging information.
2. **WPF / GUI STA Threading**: Any script that uses WPF (`PresentationFramework`) must check its apartment state on launch. If not in Single-Threaded Apartment (`STA`) mode, it must automatically re-launch itself in STA mode:
   ```powershell
   if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
       $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
       Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
       exit
   }
   ```
3. **Standardized Logging**:
   - For background / automation scripts: Import and use `Modules/Logging/Logging.psm1` (writes to `C:\Logs` with `Write-ScriptLog`).
   - For standalone Win32 / Intune packagers: Write timestamped diagnostic logs to `C:\Log\<AppName>\` (`Install.log`, `Detection.log`, `Uninstall.log`, `Packager.log`). Standardize fallback to `%LOCALAPPDATA%\Log\<AppName>\` if `C:\Log` cannot be created by standard users.
4. **PSScriptAnalyzer Compliance**: Code must pass linting defined in `PSScriptAnalyzerSettings.psd1`:
   - Use approved PowerShell verbs (`Get-`, `Set-`, `Install-`, `Remove-`, `New-`, etc.).
   - Use singular nouns for custom functions.
   - Avoid cmdlet aliases (use `Get-ChildItem` instead of `ls`/`dir`, `Copy-Item` instead of `cp`/`copy`).
   - Wrap core execution logic in `try { ... } catch { ... }` blocks with descriptive error logs.

---

## 4. Standardized README.md Walkthrough Structure

Every tool directory must contain a comprehensive `README.md` following this exact section hierarchy:

1. **Header**: `# <Tool Name> Guide`
2. **`## Overview`**: High-level summary of what the script does and the operational problem it solves.
3. **`### Key Features`**: Bulleted list of capabilities, supported OS/client targets, and mechanics.
4. **`## Prerequisites`**:
   - OS Support (e.g. Windows 10 1909+ / Windows 11)
   - PowerShell version (Windows PowerShell 5.1 / PowerShell Core 7+)
   - Permissions (e.g. Local Administrator rights required)
   - Dependencies (e.g. .NET assemblies, external tools)
5. **`## Walkthrough & Usage Guide`**:
   - `### 1. Step-by-Step Instructions`: Numbered walkthrough explaining how to launch, configure, and execute.
   - For Intune tools:
     - `### 2. Intune Portal Upload Parameters`: Markdown table with Install command, Uninstall command, Install behavior (User/System), Return codes, Detection rules, Architecture, etc.
     - `### 3. Step-by-Step Intune Admin Center Setup`: Walkthrough of the Intune creation wizard.
   - `### Logging & Diagnostics`: Exact table or list of log files written to `C:\Log\<AppName>\` or `C:\Logs\`.
   - `### How Detection & Installation Work`: Description of detection mechanics (manifest files, exit codes `0` vs `1`).
6. **`## Fast Execute`**:
   - **For Single-File Remote Scripts**: Provide short URL and GitHub transparent command:
     ```markdown
     > [!TIP]
     > **Short Branded URL (phnx.it):**
     > ```powershell
     > irm phnx.it/<CODE> | iex
     > ```
     > **Full Transparent GitHub Command:**
     > ```powershell
     > iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/<Category>/<ToolName>/<ToolName>.ps1")
     > ```
     ```
   - **For Multi-File Bundled Packages / GUI Studios**:
     ```markdown
     > [!WARNING]
     > **Bundled Package Notice:**
     > This tool is a multi-file package and cannot be executed in-memory via `irm | iex`. Please download or clone the repository locally before executing.
     >
     > [!TIP]
     > **Run locally in PowerShell (as Administrator):**
     > Double-click `Start-Gui.bat` or run:
     > ```powershell
     > powershell.exe -ExecutionPolicy Bypass -File .\<ToolName>.ps1
     > ```
     ```

---

## 5. Helper Launchers & Git Hygiene

1. **1-Click Batch Launcher (`Start-Gui.bat`)**:
   For tools featuring a GUI, provide a batch launcher in the tool folder:
   ```bat
   @echo off
   title <Tool Name>
   echo ============================================================
   echo Launching <Tool Name> GUI...
   echo ============================================================
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0<ToolName>.ps1"
   if %ERRORLEVEL% NEQ 0 (
       echo.
       echo Script closed with exit code %ERRORLEVEL%. Press any key to exit.
       pause
   )
   ```
2. **Local `.gitignore`**:
   Add a `.gitignore` inside the tool folder to ignore compilation outputs:
   ```gitignore
   Output/
   *.intunewin
   ```

---

## 6. Repository Catalog & Cryptographic Baselining

1. **Update Repository Catalog**: Whenever a new script is added, register it in `PowerShell/README.md` under the repository directory structure or category description.
2. **Mandatory Baseline Synchronization**: Whenever an agent creates, modifies, or refactors scripts in `PowerShell`, the agent MUST execute `npm run sync:baselines` in the `Helper` repository. This calculates SHA-256 digests and synchronizes Cloudflare D1 so the script immediately displays as `Verified` in the portal without manual user intervention.
