# PowerShell Repository (`jamesapf-hub/PowerShell` — `scripts.pheneox.net`)

- **Skill**: Load [powershell-intune-packager](.agents/skills/powershell-intune-packager/SKILL.md) before authoring, packaging, or documenting scripts.
- **Standards (Strictly Scoped to `PowerShell/`)**:
  - **Wrapped Folder**: `Scripts/<Category>/<ToolName>/<ToolName>.ps1` + `README.md` + `.gitignore` (`Output/`, `*.intunewin`) + optional `Start-Gui.bat`. Never introduce web/Worker code here or `irm | iex` banners into `Helper`/`GuidesAndTips`.
  - **WPF STA & Logging**: Verify `ApartmentState -eq 'STA'` on `PresentationFramework` scripts; write UK-timestamped (`dd/MM/yy HH:mm:ss`) logs to `C:\Logs\<ToolName>\` (`%LOCALAPPDATA%\Logs\<ToolName>\` fallback).
  - **Intune Output `.txt` Guide**: Every Intune Win32 packager must export an instructions `.txt` file matching the 5-step Intune upload wizard (`STEP 1`–`STEP 5`) with a copy-pasteable `Description` box containing the summary, `PACKAGE DETAILS`, `WORKSTATION LOG LOCATION: C:\Logs\<PackageFolder>\`, and `LOG FILES IN THIS FOLDER` (see [references/readme-template.md](.agents/skills/powershell-intune-packager/references/readme-template.md)).
  - **PSScriptAnalyzer**: Approved verbs, singular nouns, zero aliases, `try/catch` error handling (`PSScriptAnalyzerSettings.psd1`).
- **Mandatory Baseline Sync**: Update `PowerShell/README.md` if adding/renaming a tool, then run `npm run sync:baselines` inside `C:\Users\james\OneDrive\Documents\GitHub\Helper`.
