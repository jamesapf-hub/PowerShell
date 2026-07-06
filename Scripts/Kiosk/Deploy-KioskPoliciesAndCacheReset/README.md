# Deploy Kiosk Policies and Cache Reset

---

## Overview

This script deploys native Microsoft Kiosk and Windows App group policies. It also purges local cached credentials on user logon to prevent mid-session Azure Virtual Desktop (AVD) pop-up overlaps.

---

## Prerequisites

* **OS Support:** Windows 10 / 11 / Windows Server
* **PowerShell:** Windows PowerShell 5.1 or PowerShell Core 7+
* **Permissions:** Local Administrator rights required.

---

## Walkthrough & Usage Guide

1. Run the script from an Elevated PowerShell console:
   ```powershell
   .\Deploy-KioskPoliciesAndCacheReset.ps1
   ```
