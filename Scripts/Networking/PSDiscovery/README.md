# Network switches discovery and port scanner module

---

## Overview

Queries local and remote network switches using SSH and SNMP discovery protocols. Translates active MAC tables, maps virtual LANs (VLANs), and audits port configuration maps dynamically.

> [!NOTE]
> **Switch Audit Log Location:** `C:\Logs\PSDiscovery\get-switchinfo.log` (or `$env:SystemDrive\Logs\PSDiscovery\get-switchinfo.log`)

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Authentication:** Access credentials for the network switches (SSH username/password or SNMP read-only community strings).

---

## Walkthrough & Usage Guide

1. Open PowerShell and run the main network switch discovery script:
   ```powershell
   .\get-switchinfo.ps1 -SwitchIp "192.168.1.1" -Community "public"
   ```
