KMS_Error_Fix.ps1 can be ran manually (Elevated PowerShell Prompt) or through Configuration Manager (SCCM)

**The only line you need to edit on the script is $LogDir:**<p>
$LogDir = "" # Example C:\CompanyXYZ\logs\KMSFix

**SCCM Detection Method File System:**

**Setting Type:** File System<p>
**Type:** File<p>
**Path:** Path to the Directory you choose in the Script for $LogDir<p>
**File** or Folder name:** Successful.log<p>
**Uncheck** This file or folder is associated with a 32-bit application on 64-bit systems
