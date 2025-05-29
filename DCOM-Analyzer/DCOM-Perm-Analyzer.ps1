# DCOM Enumeration and Permission Fix Script
# Requires Administrator privileges
# All actions are logged for review

# Define the output log file
$logFile = "C:\DCOM_Comprehensive_Log.txt"
if (Test-Path $logFile) { Remove-Item $logFile }

Write-Output "Starting DCOM Analysis and Permission Fix..." | Tee-Object -FilePath $logFile

# Function to log messages to both console and file
function Log {
    param ([string]$message)
    Write-Output $message | Tee-Object -FilePath $logFile -Append
}

# Function to get registry value safely
function Get-RegistryValue {
    param (
        [string]$keyPath,
        [string]$valueName
    )
    try {
        Get-ItemProperty -Path $keyPath -Name $valueName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $valueName
    } catch {
        return $null
    }
}

# Step 1: Gather DCOM CLSID errors from Event Viewer
Log "===== Step 1: Fetching DCOM CLSID Errors from Event Viewer ====="
$dcomErrors = Get-WinEvent -LogName System | Where-Object { $_.Id -eq 10016 } | ForEach-Object {
    $message = $_.Message
    $clsidMatch = ($message -match 'CLSID\s+({.*?})') ? $matches[1] : $null
    $appIdMatch = ($message -match 'AppID\s+({.*?})') ? $matches[1] : $null
    if ($clsidMatch) {
        [PSCustomObject]@{
            CLSID = $clsidMatch
            AppID = $appIdMatch
            Message = $message
        }
    }
}

if (-not $dcomErrors) {
    Log "No DCOM CLSID Errors Found!"
    exit
}

# Step 2: Map CLSIDs and AppIDs to Applications and Services
Log "===== Step 2: Mapping CLSIDs and AppIDs to Applications ====="
foreach ($error in $dcomErrors) {
    $clsid = $error.CLSID
    $appid = $error.AppID
    Log "Processing CLSID: $clsid"
    Log "Processing AppID: $appid"

    # Check CLSID in the registry
    $clsidKey = "HKCR:\CLSID\$clsid"
    $appName = Get-RegistryValue -keyPath $clsidKey -valueName "(default)"
    if ($appName) {
        Log "CLSID $clsid maps to Application: $appName"
    } else {
        Log "CLSID $clsid not found in Registry!"
    }

    # Check AppID in the registry (optional)
    if ($appid) {
        $appidKey = "HKCR:\AppID\$appid"
        $appDesc = Get-RegistryValue -keyPath $appidKey -valueName "(default)"
        if ($appDesc) {
            Log "AppID $appid maps to Application: $appDesc"
        } else {
            Log "AppID $appid not found in Registry!"
        }
    }
}

# Step 3: Audit DCOM Permissions for Each CLSID
Log "===== Step 3: Auditing DCOM Permissions ====="
foreach ($error in $dcomErrors) {
    $clsid = $error.CLSID
    Log "Auditing Permissions for CLSID: $clsid"

    try {
        $securityDescriptor = (New-Object -ComObject WbemScripting.SWbemLocator).ConnectServer("localhost", "root\cimv2").Get("Win32_DCOMApplicationSetting.AppID='$clsid'")
        $permissions = $securityDescriptor.GetAccessSecurityDescriptor().Descriptor
        Log "Permissions for CLSID $clsid: $permissions"
    } catch {
        Log "Failed to retrieve permissions for CLSID: $clsid"
    }
}

# Step 4: Fix Permissions for Each CLSID
# Uncomment the following section if you want to automatically fix permissions
<#
Log "===== Step 4: Fixing DCOM Permissions ====="
foreach ($error in $dcomErrors) {
    $clsid = $error.CLSID
    Log "Fixing Permissions for CLSID: $clsid"

    try {
        # Set permissions for SYSTEM and Administrators
        $appObj = (New-Object -ComObject WbemScripting.SWbemLocator).ConnectServer("localhost", "root\cimv2").Get("Win32_DCOMApplicationSetting.AppID='$clsid'")
        $securityDescriptor = $appObj.GetAccessSecurityDescriptor()
        $securityDescriptor.Descriptor.DACL += New-Object System.Management.ManagementBaseObject("Win32_ACE")
        $appObj.PutAccessSecurityDescriptor($securityDescriptor)
        Log "Permissions fixed for CLSID: $clsid"
    } catch {
        Log "Failed to fix permissions for CLSID: $clsid"
    }
}
#>

# Step 5: Log Instructions for Manual Remediation
Log "===== Step 5: Instructions for Manual Remediation ====="
Log "1. Use the details in this log file to identify the CLSIDs and corresponding applications."
Log "2. Open 'dcomcnfg' to adjust permissions for the problematic applications:"
Log "   - Open Component Services: Win + R > type 'dcomcnfg' > OK"
Log "   - Navigate to Component Services > Computers > My Computer > DCOM Config."
Log "   - Find the application by name or CLSID, right-click > Properties > Security Tab."
Log "   - Update Launch and Activation Permissions for SYSTEM, Administrators, or other required accounts."
Log "3. Restart the affected services or system if needed."
Log "Detailed results saved to $logFile"

# Final output
Log "DCOM Analysis Complete!"
Write-Host "Analysis complete. Results saved to $logFile" -ForegroundColor Green