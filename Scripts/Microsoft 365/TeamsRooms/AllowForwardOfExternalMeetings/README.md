# ExternalTeamsRoom Guide

## Overview
Securely enables a Microsoft Teams Room (MTR) mailbox to accept forwarded external meeting invites while blocking direct bookings from outside the organization. The successful configuration relies on a two-step process that utilizes Mail Flow Security and Calendar Processing.

### Key Features
* **Security Check:** Blocks all unauthenticated external senders from booking or emailing the room directly by requiring sender authentication.
* **Functionality Check:** Allows the Calendar Attendant to correctly interpret external meeting data contained within a forwarded invite.
* **Link Retention:** Prevents the meeting body from being stripped, ensuring the Teams/Zoom/Webex Join URL is preserved for the one-touch Join button on the console.
* **Logging:** Records all actions, successes, and errors to a local log file for auditing.

> [!NOTE]
> **Log File Location:** `$env:SystemDrive\Logs\ExternalTeamsRoom\ExternalTeamsRoom.log`

## Prerequisites
OS Support: Windows 10 / 11 or Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Exchange Online Administrator
Execution Policy: Bypass or RemoteSigned
Dependencies: ExchangeOnlineManagement module

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions
1. Open an elevated PowerShell prompt.
2. Ensure you have the Exchange Online Management module installed (`Install-Module -Name ExchangeOnlineManagement`).
3. Execute the script `ExternalTeamsRoom.ps1` and specify the `-RoomEmail` parameter. 
   For example:
   ```powershell
   .\ExternalTeamsRoom.ps1 -RoomEmail "room@contoso.com"
   ```
4. The script will check for the Exchange Online module and prompt you to authenticate to Exchange Online.
5. It will configure Mail Flow Security to reject incoming mail from unauthenticated senders (`RequireSenderAuthenticationEnabled $True`).
6. It will configure Calendar Processing to process external meetings (`ProcessExternalMeetingMessages $True`) and retain the meeting details (`DeleteComments $False`).
7. Once completed, it will automatically disconnect your Exchange Online session and log the completion.

### 2. Logging & Outputs
The script provides color-coded console output and writes to a local log file at `$env:SystemDrive\Logs\ExternalTeamsRoom\ExternalTeamsRoom.log`.
- **Cyan (INFO):** Standard process steps.
- **Green (SUCCESS):** Configuration applied successfully.
- **Red (ERROR):** Any failure to connect or configure the mailbox settings.

Verify success by checking the console output or log file for "Calendar processing successfully configured" and "Security restriction applied successfully".


## Fast Execute
> [!TIP]
> **Run Directly in PowerShell (as Administrator):**
> You can download and execute this script instantly without saving the file locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Microsoft 365/TeamsRooms/AllowForwardOfExternalMeetings/ExternalTeamsRoom.ps1")
> ```
