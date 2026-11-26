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
- Standardize all WPF and WinForms GUIs on the **Uniform Obsidian Teal + Figtree Standard** (see [references/readme-template.md](./references/readme-template.md) for canonical helpers & starter XAML):
  - **Window Size:** Compute dynamically from `[System.Windows.SystemParameters]::WorkArea` (WPF) or `Screen.PrimaryScreen.WorkingArea` (WinForms): `70%` width (clamped `860–1440`) × `80%` height (clamped `640–900`), bounded by `WorkArea - 24`, centered via `Left`/`Top` (`WindowStartupLocation="Manual"` / `StartPosition="Manual"`), resizable.
  - **Tool-Only Headers:** Window titlebars and headers show strictly the tool name & version (**never** prefix with `Pheneox`).
  - **Palette & Shell:** Radial navy canvas (`#0F1F3D -> #091122 -> #05070C`), `#B3070D19` header/footer bars, `1px` crisp hairlines (`#1FFFFFFF` / `#26354D`), `#660D192E` translucent surfaces, `#99070D1A` inputs, `#CC050A14` selectable monospace consoles (`IsReadOnly="True"`), `#2DD4BF` primary CTA (`#5EEAD4` hover, `#14B8A6` pressed, `#042F2E` foreground), `#14FFFFFF` secondary buttons (`#E2E8F0` text), `#26EF4444`/`#FCA5A5` danger buttons, and dark `DataGrid` styling (`#0B1528` header, `#091122`/`#0C172B` rows, `#1E3A5F` selection). Zero legacy purple/indigo/orange/Catppuccin/light themes, zero floating status dots, zero nested rounded boxes; 4-band shell (`Header -> Optional 1-Row Tabs/Toolbar -> Flush Workspace -> Footer`).
  - **Single-Row Footer:** Always use a 2-column `<Grid>` footer (`Column 0 Width="*"` truncated Left Hint / Log Path | `Column 1 Width="Auto"` horizontal Right Action Cluster) — never wrap or stack footer buttons.
  - **Typography:** Load static `Figtree` TTFs (`Figtree-Regular.ttf`, `Figtree-SemiBold.ttf`, `Figtree-Bold.ttf`) from `assets/fonts/` (auto-cached into `%LOCALAPPDATA%\PowerShellGui\fonts\` on `irm | iex`), applied via `#Figtree, Segoe UI Variable, Segoe UI` (WPF) or script-scoped `$Script:UiFontCollection` (`PrivateFontCollection` kept at `$Script:` scope so GC does not free GDI+ handles in WinForms) + `Cascadia Mono, Consolas` for logs/paths/numbers. Strip hardcoded `FontFamily="Segoe UI*"` / `"Calibri"` / `"Arial"` from non-monospace WPF XAML so `$Window.FontFamily` inherits cleanly.
  - **Three Archetypes:** (1) *Intune Packagers*: 2-column split (Inputs left · Artifacts + Console right, zero scrolling); (2) *M365 Consoles*: single-row horizontal section tabs (`Overview`, `Licences`, `User Directory`, `Activity Log`, `Connection`) for uniform height; (3) *Visual Pickers & Cleaners*: fluid 5-column `UniformGrid` without unnecessary scrollbars.
- Keep Intune packager WPF GUIs minimal (`AddPrinter.ps1` pattern): include only essential payload/output inputs, selectable log console, and Build/Open buttons. Auto-download and Authenticode-verify `IntuneWinAppUtil.exe` silently without exposing tool-path pickers or Intune command boxes in the UI (commands belong in the exported `.txt` guide).
- For endpoint `Install`/`Uninstall` scripts invoked by Intune IME (32-bit default host), include a 64-bit `SysNative` relaunch guard before DISM/feature or 64-bit registry operations, and **never** use `Write-Output` inside `Write-Log` (prevents pipeline return-value pollution such as false `3010` exit codes).
- Log in UK format (`dd/MM/yy HH:mm:ss`) to `C:\Logs\<ToolName>\` (e.g. `C:\Logs\Printer_<PrinterName>\`, `C:\Logs\<AppName>\` with `Install.log`, `Detection.log`, `Uninstall.log`, and `$env:LOCALAPPDATA\Logs\<ToolName>\` fallback).
- For any Intune Win32 (`.intunewin`) packager, always export a `.txt` deployment guide alongside the `.intunewin` and detection `.ps1` script that strictly follows the 5-step Intune upload wizard (`STEP 1: SELECT APP PACKAGE FILE`, `STEP 2: APP INFORMATION` with the copy-pasteable `Description` box listing `C:\Logs\<PackageFolder>\` and all logs inside it, `STEP 3: PROGRAM`, `STEP 4: REQUIREMENTS`, `STEP 5: DETECTION RULES`) defined in [references/readme-template.md](./references/readme-template.md).
- Pass `PSScriptAnalyzerSettings.psd1`: approved verbs, singular nouns, zero cmdlet aliases (`ls`/`%`/`?`), and `try/catch` error logging.

## 3. Mandatory Post-Change Sync
1. Update `PowerShell/README.md` if adding/renaming a tool.
2. Run `npm run sync:baselines` in `C:\Users\james\OneDrive\Documents\GitHub\Helper` to update `metadata.json` and D1 SHA-256 `script_baselines`.
