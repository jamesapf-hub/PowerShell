# Intune Printer Deployment Packager

---

## Overview

A premium packaging and installation suite to deploy printers silently via Microsoft Intune. It supports TCP/IP network ports, custom vendor driver installation, and configuration profile packaging.

> [!NOTE]
> **Log File Location:** `C:\Logs\AddPrinter\<PrinterName>\Install-Printer.log` (or `$env:SystemDrive\Logs\AddPrinter\<PrinterName>\Install-Printer.log`)

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required to register drivers and port configurations.

---

## Walkthrough & Usage Guide

1. Launch the local printer packager assistant GUI by running the batch file `launchcontrol.bat`.
2. Enter the target IP Address, Driver Name, and Port details.
3. Package the printer installer payload, and upload it as a Win32 App (.intunewin) package to Microsoft Intune.
