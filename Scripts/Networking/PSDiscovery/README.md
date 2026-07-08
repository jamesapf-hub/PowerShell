# Switch Information & Port Discovery Extractor (`PSDiscovery`)

> [!NOTE]
> **Switch Audit Log Location:** `C:\Logs\PSDiscovery\get-switchinfo.log` (or `$env:SystemDrive\Logs\PSDiscovery\get-switchinfo.log`)

This utility discovers switch port details passively using CDP/LLDP packet capturing and queries switch interfaces via SNMP. It includes a built-in installer for local, offline execution.

## Features
- **Passive Switch Port Discovery**: Captures CDP and LLDP packets to discover what switch, port, and VLAN your computer is physically plugged into (requires Administrator privileges).
- **Active SNMP Switch Auditing**: Queries any network switch via SNMP v2c to map interface indexes, descriptions, operational statuses, and active VLAN assignments (does not require Administrator privileges).
- **Fast Execute & Auto-Installer**: Run a single command directly from GitHub to execute immediately, download local copies of scripts and dependencies (including `PSDiscoveryProtocol` and `Lextm.SharpSnmpLib` DLL), and register shortcut functions in your PowerShell Profile.

---

## Fast Execute & Installation

> [!TIP]
> **Run Directly in PowerShell (to Install Locally):**
> Execute the following command in PowerShell. It will automatically run the installer, download all scripts and offline dependencies, and register shortcuts:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Networking/PSDiscovery/get-switchinfo.ps1")
> ```

---

## Offline Usage
Once installed, open a new PowerShell window and run:

### 1. Unified Interactive Menu
Simply type:
```powershell
get-switchinfo
```
Choose option `1` to run a real SNMP audit, `2` to run an LLDP/CDP port capture, or `3` to reinstall/update the tool locally.

### 2. Active SNMP Switch Audits
Query a switch's interface VLAN assignments:
```powershell
get-switchinfo -IPAddress 192.168.1.1 -Community public
```

### 3. Passive CDP/LLDP Capture
Capture and save physical link info directly to local CSV logs:
```powershell
Get-SwitchPortInfo
```

---

## Prerequisites & Requirements
- **PowerShell**: Windows PowerShell `v5.1` or PowerShell Core `v7.0`+
- **Privileges**: 
  - **Standard User**: Can query switches via SNMP (`get-switchinfo -IPAddress <IP>`).
  - **Administrator**: Required for passive packet capture (`Get-SwitchPortInfo` / Option 2) to hook network adapter trace sessions.
- **Port Access**: Local port access to switch interfaces (SNMP port `161` UDP) must be open.
