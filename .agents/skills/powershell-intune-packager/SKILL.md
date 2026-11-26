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
- Standardize all WPF and WinForms GUIs on the **Uniform Obsidian Teal / Royal Blue + Figtree + Painted Win11 Chrome Standard** (see [references/readme-template.md](./references/readme-template.md) for canonical helpers & starter XAML):
  - **Window Size:** Compute dynamically from `[System.Windows.SystemParameters]::WorkArea` (WPF) or `Screen.PrimaryScreen.WorkingArea` (WinForms): `70%` width (clamped `860–1440`) × `80%` height (clamped `640–900`), bounded by `WorkArea - 24`, centered via `Left`/`Top` (`WindowStartupLocation="Manual"` / `StartPosition="Manual"`), resizable.
  - **Tool-Only Headers:** Window titlebars and headers show strictly the tool name & version (**never** prefix with `Pheneox`), plus a header `BtnThemeToggle` (`TxtThemeLabel` displaying `Theme: System (Dark/Light)` or `Theme: Manual (Dark/Light)`).
  - **Painted Win11 Chrome:** Include guarded `Win11ChromeHelper` (`DwmSetWindowAttribute` `20`/`33`/`34`/`35`/`36` matching Frame's dark `#0F1626` caption / `#F8FAFC` text / `#1E293B` border and light `#FFFFFF` caption / `#0F172A` text / `#CBD5E1` border, wrapped in `try/catch` for Windows 10) wired on `SourceInitialized` (WPF) or `HandleCreated` (WinForms).
  - **Automatic System Light/Dark Mode:** Read `HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme` via `Get-SystemIsDarkTheme`, bind WPF colors via `{DynamicResource ...}` (with `Background="#091122"` fallback on `<Window>`), and swap frozen brush resources cleanly via `Set-StandardGuiTheme` (assigning `$Window.Background = $canvasBrush` directly — **never** `DynamicResource BrushCanvas` on `<Window>`!). Dark mode uses **Obsidian Teal** (`#2DD4BF` CTA, `#5EEAD4` hover, `#14B8A6` pressed, `#042F2E` foreground) on a radial navy canvas (`#0F1F3D -> #091122 -> #05070C`), `#CC070D19` header/footer bars, `1px` hairlines (`#1CFFFFFF` / `#26354D`), `#B30D192E` cards, `#C7070D1A` inputs, and `#E6050A14` selectable monospace consoles (`IsReadOnly="True"`). Light mode uses **Royal Blue** (`#2563EB` CTA, `#1D4ED8` hover, `#1E40AF` pressed, `#FFFFFF` foreground) on `#F8F9FF` canvas with `#FFFFFF` cards, `#CBD5E1` borders, `#F8FAFC` inputs, and `#F1F5F9` consoles. Wire `BtnThemeToggle` (`Click`) and `[Microsoft.Win32.SystemEvents]::add_UserPreferenceChanged` (cleaning up the handler on `Closed`). Zero legacy purple/indigo/orange/Catppuccin themes, zero floating status dots, zero nested rounded boxes; 4-band shell (`Header -> Optional 1-Row Tabs/Toolbar -> Flush Workspace -> Footer`).
  - **Single-Row Footer:** Always use a 2-column `<Grid>` footer (`Column 0 Width="*"` truncated Left Hint / Log Path | `Column 1 Width="Auto"` horizontal Right Action Cluster) — never wrap or stack footer buttons.
  - **Typography:** Load static `Figtree` TTFs (`Figtree-Regular.ttf`, `Figtree-SemiBold.ttf`, `Figtree-Bold.ttf`) from `assets/fonts/` (auto-cached into `%LOCALAPPDATA%\PowerShellGui\fonts\` on `irm | iex`), applied via `#Figtree, Segoe UI Variable, Segoe UI` (WPF) or script-scoped `$Script:UiFontCollection` (`PrivateFontCollection` kept at `$Script:` scope so GC does not free GDI+ handles in WinForms) + `Cascadia Mono, Consolas` for logs/paths/numbers. Strip hardcoded `FontFamily="Segoe UI*"` / `"Calibri"` / `"Arial"` from non-monospace WPF XAML so `$Window.FontFamily` inherits cleanly.
  - **Three Archetypes:** (1) *Intune Packagers*: 2-column split (Inputs left · Artifacts + Console right, zero scrolling); (2) *M365 Consoles*: single-row horizontal section tabs (`Overview`, `Licences`, `User Directory`, `Activity Log`, `Connection`) for uniform height; (3) *Visual Pickers & Cleaners*: fluid 5-column `UniformGrid` without unnecessary scrollbars.
- Keep Intune packager WPF GUIs minimal (`AddPrinter.ps1` pattern): include only essential payload/output inputs, selectable log console, and Build/Open buttons. Auto-download and Authenticode-verify `IntuneWinAppUtil.exe` silently without exposing tool-path pickers or Intune command boxes in the UI (commands belong in the exported `.txt` guide).
- For endpoint `Install`/`Uninstall` scripts invoked by Intune IME (32-bit default host), include a 64-bit `SysNative` relaunch guard before DISM/feature or 64-bit registry operations, and **never** use `Write-Output` inside `Write-Log` (prevents pipeline return-value pollution such as false `3010` exit codes).
- Log in UK format (`dd/MM/yy HH:mm:ss`) to `C:\Logs\<ToolName>\` (e.g. `C:\Logs\Printer_<PrinterName>\`, `C:\Logs\<AppName>\` with `Install.log`, `Detection.log`, `Uninstall.log`, and `$env:LOCALAPPDATA\Logs\<ToolName>\` fallback).
- For any Intune Win32 (`.intunewin`) packager, always export a `.txt` deployment guide alongside the `.intunewin` and detection `.ps1` script that strictly follows the 5-step Intune upload wizard (`STEP 1: SELECT APP PACKAGE FILE`, `STEP 2: APP INFORMATION` with the copy-pasteable `Description` box listing `C:\Logs\<PackageFolder>\` and all logs inside it, `STEP 3: PROGRAM`, `STEP 4: REQUIREMENTS`, `STEP 5: DETECTION RULES`) defined in [references/readme-template.md](./references/readme-template.md).
- **UTF-8 Without BOM (`irm | iex` Safety):** Save every `.ps1` file as **UTF-8 without BOM** (`[System.Text.UTF8Encoding]::new($false)` — never PS 5.1 `Set-Content -Encoding UTF8`), because `Invoke-RestMethod` preserves a leading `U+FEFF` BOM when piping into `iex`, breaking `<#` comment-based help headers.
- Pass `PSScriptAnalyzerSettings.psd1`: approved verbs, singular nouns, zero cmdlet aliases (`ls`/`%`/`?`), and `try/catch` error logging.

## 3. Mandatory Post-Change Sync
1. Update `PowerShell/README.md` if adding/renaming a tool.
2. Run `npm run sync:baselines` in `C:\Users\james\OneDrive\Documents\GitHub\Helper` to update `metadata.json` and D1 SHA-256 `script_baselines`.
