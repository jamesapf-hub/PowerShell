# Install Datto RMM Agent

---

## Overview

Downloads and silently installs the Datto RMM Agent using site credentials. This script is sanitized to use variables for site details.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.
* **Configuration:** Edit the script variables `$Platform` and `$SiteID` to match your Datto client site setup before execution.

---

## Walkthrough & Usage Guide

1. Open the script and set your Datto details:
   ```powershell
   $Platform="YOUR_DATTO_PLATFORM"
   $SiteID="YOUR_DATTO_SITE_ID"
   ```
2. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Install-DattoRmmAgent.ps1
   ```
