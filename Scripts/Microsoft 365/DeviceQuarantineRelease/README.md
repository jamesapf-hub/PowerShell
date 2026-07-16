# Device Quarantine Release Guide

## Overview
The `DeviceQuarantineReleaseApp.ps1` script launches a native Windows WPF desktop assistant designed for Exchange Online administration. It provides a simple, graphical interface to search tenant mailboxes, retrieve ActiveSync mobile device metadata, inspect their current quarantine status, and instantly release selected devices to the Allowed list.

### Key Features
* **ActiveSync Status Inspection:** Retrieves all mobile devices associated with a mailbox, showing Device Model, OS, Client Type, Access State, and Quarantine reason.
* **One-Click Device Release:** Automatically updates user CAS mailbox configurations to allow-list specific device identifiers.
* **Auto-Module Provisioning:** Verifies system modules at launch and auto-installs the `ExchangeOnlineManagement` package if not found.
* **Live Activity Telemetry:** Includes a scrolling logs viewport at the bottom to track cmdlet execution states and session status.

## Prerequisites
OS Support: Windows 10 / 11 (requires WPF rendering libraries)
PowerShell: PowerShell 5.1 or PowerShell 7+
Permissions: Helpdesk Administrator, Exchange Recipient Administrator, or Exchange Administrator role in Exchange Online.
Dependencies: `ExchangeOnlineManagement` module (script will automatically prompt/install if missing).

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### How to Launch the GUI
Run the script using your terminal (bypassing execution policies):
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\DeviceQuarantineReleaseApp.ps1
```

#### Using the App
1. **Connect:** Click the **Connect** button in the top-right corner. Complete the Microsoft 365 modern authentication popup.
2. **Load Directory:** Click **Load Mailboxes** to fetch licensed mailbox users.
3. **Filter & Search:** Use the search bar to locate specific display names or user emails.
4. **Inspect Devices:** Select a user from the left-side list; their registered mobile devices will load in the right-side details panel.
5. **Release Device:** If a device's access state displays as **Quarantined**, select the device and click the green **Release Device** button.
6. **Disconnect:** Click **Disconnect** when done to securely close the Exchange Online PowerShell session.

### 2. Logging & Outputs
* **GUI Live Feed:** Underlying operations (including remote connection, directory fetching, CAS mailbox updates, and disconnect handshakes) write telemetry entries directly to the console area at the bottom.

## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Microsoft 365/DeviceQuarantineRelease/DeviceQuarantineReleaseApp.ps1")
> ```
