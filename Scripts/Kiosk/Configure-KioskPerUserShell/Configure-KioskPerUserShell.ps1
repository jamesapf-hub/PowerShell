Param(
    [string]$KioskUser = "User",
    [string]$ShellPath = "C:\Program Files\Wyse\WyseEasySetup\WyseEasySetupShell.exe"
)

# Define log paths and targets
$LogFolder = "$env:SystemDrive\Logs\Configure-KioskPerUserShell"
$LogFile   = "$LogFolder\Wyse_Shell_Lock.txt"
$RegPath   = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

# 1. Ensure log folder exists
if (-not (Test-Path $LogFolder)) { 
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null 
}

# Write to log file and output directly to console for real-time visibility
function Write-Log ($Msg) { 
    $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "[ $TimeStamp ] $Msg" | Out-File $LogFile -Append 
    Write-Host "[*] $Msg"
}

# 1a. Validate Administrator Elevation
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "ERROR: This script must be run as an Administrator (Elevated PowerShell session)."
    Write-Host "Please open PowerShell as Administrator and run the script again." -ForegroundColor Red
    exit 1
}

# 1b. C# Token Privilege Elevation Helper (to bypass registry DACL lockouts)
$definition = @"
using System;
using System.Runtime.InteropServices;

public class TokenPrivilegeHelper {
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool OpenProcessToken(IntPtr ProcessHandle, uint DesiredAccess, out IntPtr TokenHandle);
    
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool LookupPrivilegeValue(string lpSystemName, string lpName, out long lpLuid);
    
    [DllImport("advapi32.dll", SetLastError = true)]
    public static extern bool AdjustTokenPrivileges(IntPtr TokenHandle, bool DisableAllPrivileges, ref TokPriv1Luid NewState, int BufferLength, IntPtr PreviousState, IntPtr ReturnLength);

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct TokPriv1Luid {
        public int Count;
        public long Luid;
        public int Attr;
    }

    public static void Enable(string privilege) {
        IntPtr hToken = IntPtr.Zero;
        if (OpenProcessToken(System.Diagnostics.Process.GetCurrentProcess().Handle, 0x0020 | 0x0008, out hToken)) {
            long luid = 0;
            if (LookupPrivilegeValue(null, privilege, out luid)) {
                TokPriv1Luid tp = new TokPriv1Luid();
                tp.Count = 1;
                tp.Luid = luid;
                tp.Attr = 2; // SE_PRIVILEGE_ENABLED
                AdjustTokenPrivileges(hToken, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $definition -ErrorAction SilentlyContinue
    [TokenPrivilegeHelper]::Enable("SeTakeOwnershipPrivilege")
    [TokenPrivilegeHelper]::Enable("SeRestorePrivilege")
    Write-Log "SUCCESS: Enabled SeTakeOwnershipPrivilege and SeRestorePrivilege for current process."
} catch {
    Write-Log "WARNING: Failed to enable token privileges. Reason: $_"
}

Write-Log "=== Shell Setup (Per-User Shell Mode) Started ==="
Write-Log "Target Kiosk User: $KioskUser"
Write-Log "Target Shell Path: $ShellPath"

# 2. Revert registry permission locks on HKLM Winlogon to prevent boot failures (BSODs)
Write-Log "Checking and cleaning HKLM Winlogon ACLs..."
$RegKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

# Step 2a: Take ownership of the key (requires SeTakeOwnershipPrivilege to bypass DACL check)
try {
    # Open key requesting ONLY TakeOwnership rights
    $Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
        $RegKeyPath, 
        [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadSubTree, 
        [System.Security.AccessControl.RegistryRights]::TakeOwnership
    )
    if ($Key) {
        # Bypass the .NET internal "read-only" check using reflection
        $Type = $Key.GetType()
        $Flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        
        # Method A: Try boolean fields (mostly .NET Core / .NET 5+)
        $Field = $Type.GetField("writable", $Flags)
        if (-not $Field) { $Field = $Type.GetField("_writable", $Flags) }
        if ($Field) { $Field.SetValue($Key, $true) }
        
        # Method B: Try integer state fields (mostly .NET Framework 4.x used by PowerShell 5.1)
        $StateField = $Type.GetField("state", $Flags)
        if (-not $StateField) { $StateField = $Type.GetField("_state", $Flags) }
        if ($StateField) {
            $currentVal = $StateField.GetValue($Key)
            $newVal = $currentVal -bor 4 # STATE_WRITEABLE is 4
            $StateField.SetValue($Key, $newVal)
        }

        # Get the current owner section of the ACL
        $Acl = $Key.GetAccessControl([System.Security.AccessControl.AccessControlSections]::Owner)
        
        # Resolve the current running identity (SYSTEM or Administrator) and set it as the owner
        $CurrentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $Acl.SetOwner($CurrentIdentity)
        
        # Apply ownership change
        $Key.SetAccessControl($Acl)
        $Key.Close()
        Write-Log "SUCCESS: Took ownership of HKLM Winlogon key."
    }
} catch {
    Write-Log "WARNING: Failed to take ownership of Winlogon key. Reason: $_"
}

# Step 2b: Open the key requesting ChangePermissions rights
# (This now succeeds because the current identity is the owner, which grants WRITE_DAC bypass)
try {
    $Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
        $RegKeyPath, 
        [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadSubTree, 
        [System.Security.AccessControl.RegistryRights]::ChangePermissions
    )
    if ($Key) {
        # Bypass the .NET internal "read-only" check using reflection
        $Type = $Key.GetType()
        $Flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        
        # Method A: Try boolean fields (mostly .NET Core / .NET 5+)
        $Field = $Type.GetField("writable", $Flags)
        if (-not $Field) { $Field = $Type.GetField("_writable", $Flags) }
        if ($Field) { $Field.SetValue($Key, $true) }
        
        # Method B: Try integer state fields (mostly .NET Framework 4.x used by PowerShell 5.1)
        $StateField = $Type.GetField("state", $Flags)
        if (-not $StateField) { $StateField = $Type.GetField("_state", $Flags) }
        if ($StateField) {
            $currentVal = $StateField.GetValue($Key)
            $newVal = $currentVal -bor 4 # STATE_WRITEABLE is 4
            $StateField.SetValue($Key, $newVal)
        }

        # Get the access rules section of the ACL
        $Acl = $Key.GetAccessControl([System.Security.AccessControl.AccessControlSections]::Access)
        
        $RulesToRemove = @()
        foreach ($Rule in $Acl.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])) {
            if ($Rule.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny) {
                $RulesToRemove += $Rule
            }
        }
        
        if ($RulesToRemove.Count -gt 0) {
            Write-Log "Found $($RulesToRemove.Count) Deny rule(s). Removing..."
            foreach ($Rule in $RulesToRemove) {
                $Acl.RemoveAccessRule($Rule) | Out-Null
            }
            # Set the updated ACL
            $Key.SetAccessControl($Acl)
            Write-Log "SUCCESS: Reverted Deny rules on Winlogon registry key."
        } else {
            Write-Log "INFO: No Deny rules found on HKLM Winlogon. Key is unlocked."
        }
        $Key.Close()
    }
} catch {
    Write-Log "ERROR: Failed to clean registry key ACL. Reason: $_"
}

# 3. Restore global HKLM shell to explorer.exe
try {
    Set-ItemProperty -Path $RegPath -Name "Shell" -Value "explorer.exe" -Type String -Force -ErrorAction Stop
    Write-Log "SUCCESS: Set global (HKLM) Shell value back to explorer.exe"
} catch {
    Write-Log "ERROR: Failed to restore HKLM Shell. Reason: $_"
}

# 3a. Remove per-user shell override for current user if they are not the kiosk user
# (This ensures the Admin account immediately regains Explorer shell if it was overridden in HKCU)
if ($env:USERNAME -ne $KioskUser) {
    $CurrentHKCUWinlogon = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
    if (Get-ItemProperty -Path $CurrentHKCUWinlogon -Name "Shell" -ErrorAction SilentlyContinue) {
        try {
            Remove-ItemProperty -Path $CurrentHKCUWinlogon -Name "Shell" -Force -ErrorAction Stop
            Write-Log "SUCCESS: Cleared per-user shell override in HKCU for current user ($env:USERNAME) to restore Explorer."
        } catch {
            Write-Log "WARNING: Failed to remove HKCU shell override for $env:USERNAME. Reason: $_"
        }
    } else {
        Write-Log "INFO: No per-user shell override found in HKCU for current user ($env:USERNAME)."
    }
}

# 4. Set per-user shell for the Kiosk User
$ProfilePath = "C:\Users\$KioskUser"
if (Test-Path $ProfilePath) {
    $HivePath = "$ProfilePath\NTUSER.DAT"
    if (Test-Path $HivePath -PathType Leaf) {
        
        # Get the User's SID
        $UserObj = New-Object System.Security.Principal.NTAccount($KioskUser)
        $UserSID = $null
        try {
            $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier]).Value
        } catch {
            Write-Log "WARNING: Could not resolve SID for $KioskUser."
        }

        # Check if user hive is currently loaded
        if ($UserSID -and (Test-Path "Registry::HKEY_USERS\$UserSID")) {
            Write-Log "User $KioskUser is currently logged in. Setting shell in active hive..."
            try {
                $UserWinlogon = "Registry::HKEY_USERS\$UserSID\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
                if (-not (Test-Path $UserWinlogon)) {
                    New-Item -Path $UserWinlogon -Force | Out-Null
                }
                Set-ItemProperty -Path $UserWinlogon -Name "Shell" -Value $ShellPath -Type String -Force -ErrorAction Stop
                Write-Log "SUCCESS: Per-user shell set for active session of $KioskUser."
            } catch {
                Write-Log "ERROR: Failed to write to active user session. Reason: $_"
            }
        } else {
            # User is offline, load their registry hive
            Write-Log "User $KioskUser is offline. Loading user registry hive..."
            $TempKeyName = "TempHive_$KioskUser"
            
            & reg.exe load "HKU\$TempKeyName" "$HivePath" 2>&1 | Out-String | Write-Log
            
            if ($LASTEXITCODE -eq 0) {
                try {
                    $UserWinlogon = "Registry::HKEY_USERS\$TempKeyName\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    if (-not (Test-Path $UserWinlogon)) {
                        New-Item -Path $UserWinlogon -Force | Out-Null
                    }
                    Set-ItemProperty -Path $UserWinlogon -Name "Shell" -Value $ShellPath -Type String -Force -ErrorAction Stop
                    Write-Log "SUCCESS: Per-user shell configured in loaded hive for $KioskUser."
                } catch {
                    Write-Log "ERROR: Failed to set shell in loaded hive. Reason: $_"
                } finally {
                    # Release locks on the registry provider so we can unload
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                    & reg.exe unload "HKU\$TempKeyName" 2>&1 | Out-String | Write-Log
                    Write-Log "Unloaded registry hive for $KioskUser."
                }
            } else {
                Write-Log "ERROR: Failed to load user registry hive."
            }
        }
    } else {
        Write-Log "WARNING: NTUSER.DAT not found for $KioskUser at $HivePath."
    }
} else {
    Write-Log "WARNING: User profile folder not found at $ProfilePath. Skipping per-user shell setup."
}

Write-Log "=== Shell Setup Finished ==="
exit 0