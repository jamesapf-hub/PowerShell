# Ultimate PC CleanUp Utility Guide

---

## Overview

This is the main runner dashboard for the Ultimate PC CleanUp Utility. It acts as a central coordinator that discovers, schedules, and runs the individual sub-cleanup scripts on a local machine. It can be run in two modes:
1. **Interactive GUI Mode (WPF):** Launches a modern dark-themed dashboard showing discovered scripts, status badges (Pending/Running/Success/Failed), a progress bar, and a real-time console log.
2. **Command-Line (CLI) Mode:** Runs cleanup tasks directly in the terminal, making it ideal for RMM, Active Directory startup scripts, scheduled tasks, or command-line automation.

### Key Features
* **Auto-Discovery:** Scans the \`Scripts/\` subfolder recursively, parses \`.SYNOPSIS\` blocks, and displays titles and descriptions dynamically.
* **Modern Dark UI:** Premium WPF interface styled with clean layouts, custom progress indicators, and status badges.
* **Multithreaded Asynchronous Execution:** Runs scripts in a background runspace so the GUI remains active and responsive during operations.
* **Real-time Console Stream:** Captures standard outputs and errors from sub-processes asynchronously line-by-line.
* **Silent & Selective CLI Automation:** Run specific tasks or all tasks directly from the command line using parameter switches.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server (WPF required for GUI mode)
* **PowerShell:** Windows PowerShell 5.1 (WPF assemblies PresentationFramework and PresentationCore required)
* **Permissions:** Local Administrator rights required (triggers UAC elevation prompt automatically if run as standard user)
* **Execution Policy:** RemoteSigned or Bypass

---

## Walkthrough & Usage Guide

### 1. Interactive GUI Mode
1. Double-click \`Run-UltimateCleanUp.bat\` or run:
   \`\`\`powershell
   powershell.exe -ExecutionPolicy Bypass -File .\\Start-UltimateCleanUp.ps1 -Gui
   \`\`\`
2. The Ultimate PC CleanUp dashboard will open.
3. Review the discovered cleanup tasks on the left panel (you can check or uncheck individual scripts).
4. Click **Run Selected Tasks** or **Run All Tasks** to begin execution.
5. Monitor real-time progress, console outputs, and success status badges.

### 2. Command-Line (CLI) Mode
* **List discovered tasks:**
  \`\`\`powershell
  powershell.exe -ExecutionPolicy Bypass -File .\\Start-UltimateCleanUp.ps1 -ListTasks
  \`\`\`
* **Run all discovered tasks silently:**
  \`\`\`powershell
  powershell.exe -ExecutionPolicy Bypass -File .\\Start-UltimateCleanUp.ps1 -RunAll
  \`\`\`
* **Run specific tasks by name:**
  \`\`\`powershell
  powershell.exe -ExecutionPolicy Bypass -File .\\Start-UltimateCleanUp.ps1 -RunTasks "SystemTempCleaner, ClearWindowsUpdateCache"
  \`\`\`

---

## Command

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Start-UltimateCleanUp.ps1
```
