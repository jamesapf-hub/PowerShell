# Get Azure VM Details Guide

## Overview
> **Short Description:** Queries Azure IMDS metadata to display and log VM configuration, network IPs, subscription details, and region info.

Queries the Azure Instance Metadata Service (IMDS) at `http://169.254.169.254` from within an Azure Virtual Machine to retrieve compute details, resource group parameters, network interfaces, and public/private IP addresses.

### Key Features
* **IMDS Metadata Query:** Fetches live instance metadata directly from Azure hypervisor endpoint without requiring Azure PowerShell module logins.
* **Structured Console Summary:** Formats VM Name, Size, OS Type, Resource Group, Subscription ID, Location, IPs, Subnet, and MAC address.
* **Automated File Logging:** Automatically saves timestamped audit log files to `$env:SystemDrive\Logs\Get-AzureVmDetails`.

> [!NOTE]
> **Log File Location:** `C:\Logs\Get-AzureVmDetails\AzureVM_Metadata_[Timestamp].log` (or `$env:SystemDrive\Logs\Get-AzureVmDetails\AzureVM_Metadata_[Timestamp].log`)

## Prerequisites

**OS Support:** Windows / Linux Azure Virtual Machines  
**PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+  
**Permissions:** Standard User or Administrator context inside Azure Virtual Machine  
**Network:** Direct access to Azure IMDS link-local IP (`169.254.169.254`)

## Walkthrough & Usage Guide

### 1. Step-by-Step Instructions

1. Open PowerShell inside an Azure Virtual Machine console.
2. Run the script:
   ```powershell
   .\Get-AzureVmDetails.ps1
   ```
3. The script will query the non-routable link-local IP `http://169.254.169.254/metadata/instance` and display the formatted output.

### 2. Logging & Outputs
- **Console Output:** Colored summary table displaying all VM attributes.
- **Log Files:** Created automatically at `C:\Logs\Get-AzureVmDetails\AzureVM_Metadata_DDMMYY-HHMMSS.log`.

## Fast Execute

> [!TIP]
> **Run Directly in PowerShell:**
> You can execute this script directly inside an Azure VM without saving files locally:
> ```powershell
> iex (irm "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Azure/HelpfulTools/Get-AzureVmDetails/Get-AzureVmDetails.ps1")
> ```
