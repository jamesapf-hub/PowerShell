# PSDiscovery Guide

## Overview
This utility is designed for network engineers and IT administrators to identify physical switch port connections and query switch configurations. It operates in two modes: passive packet sniffs (CDP/LLDP) and active SNMP audits. It includes a built-in installer for local, offline bootstrapper execution.

### Key Features
* **CDP/LLDP Port Discovery:** Captures Cisco Discovery Protocol (CDP) and Link Layer Discovery Protocol (LLDP) multicast frames directly from network adapters to identify the connected switch name, port index, and VLAN.
* **Active SNMP Switch Auditing:** Queries remote switches via SNMP v2c to retrieve interface names, operational statuses (up/down), and active VLAN assignments.
* **SNMP MAC Table Port Lookup:** Automatically detects your local computer's MAC address and queries the switch's forwarding database (FDB Bridge MIB) to resolve the exact switch port you are connected to.
* **Subnet Auto-Discovery:** Automatically scans your default gateway, local ARP neighbor table, and common subnet suffixes to find SNMP-responsive switches if no IP is provided.
* **Fast Execute & Auto-Installer:** Runs directly from GitHub to execute immediately, downloads local copies of scripts and dependencies (including `PSDiscoveryProtocol` and `Lextm.SharpSnmpLib` DLL with active process-lock protection), and registers shortcut functions in your PowerShell Profile.

> [!NOTE]
> **Log File Location:** `C:\Logs\PSDiscovery\get-switchinfo.log` (or `$env:SystemDrive\Logs\PSDiscovery\get-switchinfo.log`)

## Prerequisites
OS Support: Windows 10 / 11 or Windows Server
PowerShell: Windows PowerShell 5.1 or PowerShell Core 7+
Permissions: Local Administrator rights required for CDP/LLDP capture, Standard User for SNMP query
Execution Policy: Bypass
Dependencies: NetAdapter (native Windows), PSDiscoveryProtocol (auto-installed), SharpSnmpLib (auto-installed)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

#### Running the Unified Interactive Menu
Run the tool without parameters to open the interactive menu utility:
```powershell
get-switchinfo
```
The menu has built-in loop protection (won't close on accidental Enter/invalid keys) and lets you:
* Choose **Option 1**: Run an SNMP Switch Audit. Leave the IP blank to **auto-discover** active SNMP switches on your subnet! Once selected, the script will automatically discover your active MAC address and attempt to map it to the exact switch port you are connected to.
* Choose **Option 2**: Run a passive LLDP/CDP port capture (redirects to the packet sniffer).
* Choose **Option 3**: Update/reinstall the utility locally.

#### Running Active SNMP Audits Manually
To query a switch's interface VLAN assignments and map your connected port:
```powershell
get-switchinfo -IPAddress 192.168.1.1 -Community public
```

#### Running Passive CDP/LLDP Capture Manually
To capture physical link info and save it directly to customer CSV logs:
```powershell
Get-SwitchPortInfo
```

### 2. Logging & Outputs
* **SNMP Audit Text Logs:** Detailed connection logs and audit outputs are appended to `C:\Logs\PSDiscovery\get-switchinfo.log`.
* **CDP/LLDP CSV Inventories:** Every successful CDP/LLDP discovery runs and appends to customer-specific inventories at `C:\Log\PSDiscovery\SwitchInventory_[CustomerName].csv` (defaults to `SwitchInventory.csv`).

## Command

> [!TIP]
> **Fast Execute Command (Online Mode):**
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Networking/PSDiscovery/get-switchinfo.ps1")
> ```
