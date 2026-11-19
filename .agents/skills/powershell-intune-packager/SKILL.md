---
name: powershell-intune-packager
description: >-
  Author, scaffold, package, and baseline wrapped-folder PowerShell scripts, WPF GUI tools,
  and Microsoft Intune Win32 (.intunewin) packagers in the PowerShell repository.
---

# PowerShell & Intune Win32 Packager Skill (`PowerShell/`)

- **Templates**: Read [references/readme-template.md](./references/readme-template.md) for the WPF `STA` relaunch snippet and required `README.md` / `Fast Execute` markdown structure.

## 1. Scaffolding & Wrapped Folder Pattern
- Scaffold via: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".agents\skills\powershell-intune-packager\scripts\new-wrapped-script.ps1" -Category "<Category>" -ToolName "<ToolName>" -Synopsis "<Summary>" [-Gui] [-Intune]`
- Structure: `Scripts/<Category>/<ToolName>/{<ToolName>.ps1, README.md, .gitignore, [Start-Gui.bat], [tools/]}`
- Valid Categories: `Azure`, `Clean Up Tools`, `Configuration`, `Datto`, `Intune`, `Kiosk`, `Microsoft 365`, `Networking`, `Sample`, `Software`.

## 2. Authoring, Logging, Intune `.txt` Guide & Analyzer Rules
- Include comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES` with `Script Name`, `Author: JP`, `Version`, UK `Created`/`Last Modified` `dd/MM/yy`, and `Log Path`), `[CmdletBinding(SupportsShouldProcess = $true)]`, `param()`, RMM env fallbacks (`$env:DATTO_*`), and `Start-Process -Wait -PassThru` exit-code checks.
- Verify `STA` apartment mode on all `PresentationFramework` WPF scripts.
- Log in UK format (`dd/MM/yy HH:mm:ss`) to `C:\Logs\<ToolName>\` (e.g. `C:\Logs\Printer_<PrinterName>\`, `C:\Logs\<AppName>\` with `Install.log`, `Detection.log`, `Uninstall.log`, and `$env:LOCALAPPDATA\Logs\<ToolName>\` fallback).
- For any Intune Win32 (`.intunewin`) packager, always export a `.txt` deployment guide alongside the `.intunewin` and detection `.ps1` script that strictly follows the 5-step Intune upload wizard (`STEP 1: SELECT APP PACKAGE FILE`, `STEP 2: APP INFORMATION` with the copy-pasteable `Description` box listing `C:\Logs\<PackageFolder>\` and all logs inside it, `STEP 3: PROGRAM`, `STEP 4: REQUIREMENTS`, `STEP 5: DETECTION RULES`) defined in [references/readme-template.md](./references/readme-template.md).
- Pass `PSScriptAnalyzerSettings.psd1`: approved verbs, singular nouns, zero cmdlet aliases (`ls`/`%`/`?`), and `try/catch` error logging.

## 3. Mandatory Post-Change Sync
1. Update `PowerShell/README.md` if adding/renaming a tool.
2. Run `npm run sync:baselines` in `C:\Users\james\OneDrive\Documents\GitHub\Helper` to update `metadata.json` and D1 SHA-256 `script_baselines`.
