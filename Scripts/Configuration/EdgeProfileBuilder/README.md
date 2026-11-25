# Edge Profile Builder Guide

WPF GUI utility with a visual thumbnail selector grid, bundled 9-color Edge icon pack, drag-and-drop custom image converter (multi-size `.ico`), and isolated Microsoft Edge profile shortcut generator with unique Windows 11 taskbar pinning IDs.

## Overview

`EdgeProfileBuilder.ps1` makes it easy to spin up multiple visual Microsoft Edge browser identities (e.g., per-tenant admin sessions, personal vs work profiles, or colored InPrivate workspaces) from a single dark-themed desktop interface. Tick as many icons as you want from the grid—five ticked icons create five isolated Edge profiles and five taskbar-pinnable shortcuts—rename them inline, or upload/drag-and-drop your own `.png`, `.jpg`, `.bmp`, `.gif`, or `.ico` images.

### Key Features

- **Visual Icon Selector Grid:** Displays large crisp icon cards with tick boxes, live selection counter badge, `Select All` / `Clear Selection`, and inline editable profile names.
- **Bundled 9-Color Default Pack (`assets\`):** Includes `Edge_CyberNeon`, `Edge_Gold`, `Edge_Lime`, `Edge_Magenta`, `Edge_Orange`, `Edge_Purple`, `Edge_Rainbow`, `Edge_Silver`, and `Edge_Teal`.
- **Custom Upload & Drag-and-Drop Auto-Converter:** Upload or drop `.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, or `.ico` files onto the grid; non-ICO images are automatically converted into lossless multi-size `.ico` files (`256`, `64`, `48`, `32`, `16` px) preserving transparency and aspect ratio.
- **Isolated Profiles & Taskbar Pinning:** Creates separate `--user-data-dir` folders (or `--profile-directory="Profile X"` flags), optional `--inprivate` launch mode, optional Desktop shortcut copies, and stamps a unique `System.AppUserModel.ID` onto every `.lnk` file so Windows 11 pins each shortcut independently on the taskbar.

## Prerequisites

- **OS Support:** Windows 10 (1909+) / Windows 11
- **PowerShell:** Windows PowerShell 5.1 or PowerShell 7+ (auto-relaunches in `-STA` mode if needed)
- **Browser:** Microsoft Edge (`msedge.exe`) installed in standard Program Files or LocalAppData paths
- **Permissions:** Standard user permissions (no local Administrator privileges required when writing to `%LOCALAPPDATA%\EdgeProfiles`)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Double-click `Start-Gui.bat` or run `.\EdgeProfileBuilder.ps1` from PowerShell.
2. Browse the visual icon grid loaded from the bundled `assets\` pack, or click **+ Upload Icons...** / drag-and-drop custom image files onto the grid.
3. Tick the icons you want (click any card's checkbox or click its thumbnail directly) and customize any profile name in the card's text box.
4. Choose **Isolated User-Data-Dir** (recommended for separate cookies and M365 logins) or **Standard Edge Profile Flag**, toggle **Launch in InPrivate** or **Also copy shortcuts to Desktop**, and verify the **Output Root Folder** (defaults to `%LOCALAPPDATA%\EdgeProfiles`).
5. Click **Generate N Profiles & Shortcuts**—Explorer opens the `Shortcuts\` folder automatically so you can launch or pin any `.lnk` to your taskbar.

### 2. Output Folder Structure

| Subfolder | Purpose |
| :--- | :--- |
| `<OutputRoot>\Shortcuts\` | Generated `.lnk` shortcuts stamped with unique `System.AppUserModel.ID` values |
| `<OutputRoot>\ConvertedIcons\` | Multi-size `.ico` files referenced by the shortcuts so icons never break |
| `<OutputRoot>\Profiles\Edge_<Name>\` | Isolated Edge `--user-data-dir` profile directories (when Isolated mode is selected) |

### Logging & Diagnostics

| Log Path | Purpose |
| :--- | :--- |
| `C:\Logs\EdgeProfileBuilder\EdgeProfileBuilder.log` | Primary execution, conversion, and shortcut generation log (UK `dd/MM/yy HH:mm:ss`) |
| `%LOCALAPPDATA%\Logs\EdgeProfileBuilder\EdgeProfileBuilder.log` | Automatic fallback log location if `C:\Logs` is not writable |

## Fast Execute

> [!TIP]
> **Full Transparent GitHub Command (`irm | iex`):**
> Automatically downloads and caches `EdgeProfileEngine.psm1`, `EdgeProfileBuilder.xaml`, and the 9 default `assets\` PNGs in `%LOCALAPPDATA%\EdgeProfileBuilder\cache\` when run remotely:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Configuration/EdgeProfileBuilder/EdgeProfileBuilder.ps1")
> ```
> **Run locally from a cloned folder:**
> Double-click `Start-Gui.bat` or run:
> ```powershell
> powershell.exe -ExecutionPolicy Bypass -File .\EdgeProfileBuilder.ps1
> ```
