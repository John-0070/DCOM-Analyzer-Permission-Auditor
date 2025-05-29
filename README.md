# DCOM-Analyzer-Permission-Auditor
Automated PowerShell tool to scan, analyze, and log DCOM permission errors (Event ID 10016) on Windows systems.
Designed for system administrators and security engineers to assist in identifying and resolving DCOM CLSID/AppID misconfigurations.

# Overview
DCOM (Distributed Component Object Model) errors—especially Event ID 10016—are common on Windows machines and often point to permission mismatches that affect COM-based services. This tool helps automate the tedious process of auditing and remediating those issues.

# Features
Event Log Scanning
Extracts DCOM-related Event ID 10016 entries from the System log.

CLSID & AppID Mapping
Matches CLSIDs and AppIDs to registry entries and application names.

Permission Inspection
Audits DCOM access control lists using WMI and COM interfaces.

Optional Permission Fixing
Prepares logic to fix permissions (commented out for safety).

Manual Fix Instructions
Step-by-step guide for using dcomcnfg to resolve permission issues.

Complete Logging
Outputs all results and actions to a single log file at C:\DCOM_Comprehensive_Log.txt.

# Requirements
Windows PowerShell 5.1+

Must be run with Administrator privileges

# Usage
Clone or copy the script to your system

Run in an elevated PowerShell terminal

powershell
Copy
Edit
.\dcom_audit.ps1
Review the output log

The results will be saved at:
C:\DCOM_Comprehensive_Log.txt

# Manual Remediation
If the script reports a problematic CLSID or AppID:

Press Win + R, type dcomcnfg, and press Enter

Navigate to:

nginx
Copy
Edit
Component Services > Computers > My Computer > DCOM Config
Find the application matching the CLSID or name

Right-click > Properties > Security

Update Launch and Activation Permissions
Add SYSTEM, Administrators, or appropriate service accounts

Restart the affected service or reboot

# Optional Fixing Logic
There is a commented-out section in the script that contains logic to auto-apply permissions using WMI and security descriptor manipulation.
Uncomment only after reviewing and understanding the implications.

# Limitations
Does not support bulk repair by default (safety-first design)

Some CLSIDs may not resolve correctly if the system registry is misconfigured

May require manual validation in Component Services for edge cases
