**The only line you need to edit on the script is $LogDir:**<p>
$LogDir = "" # Example C:\CompanyXYZ\logs\KMSFix

KMS_Error_Fix.ps1 can be ran manually (Elevated PowerShell Prompt) or through Configuration Manager (SCCM)

# Manual Install:<h1><p>
(Open Powershell as Administrator) Run [KMS_Error_Fix.ps1](https://github.com/cfreeman21/scripts/blob/master/KMSfix.cmd)<p>
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_Manual.png)  

# Configuration Manager (SCCM) Install:<h1><p>
[KMSfix.cmd](https://github.com/cfreeman21/scripts/blob/master/KMSfix.cmd) & [KMS_Error_Fix.ps1](https://github.com/cfreeman21/scripts/blob/master/KMSfix.cmd) need to be in same directory:
File Structure:
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_File_Structure.png)<p>
**Application Properties: (General Tab)**
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_1.png)<p>
**Application Properties: (Programs Tab)**
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_2.png)<p>
**Application Properties: (Detection Tab / Detection Rule)**
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_3.png)<p>
**Application Properties: (User Experience Tab)**
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_4.png)<p>
**Application Properties: (Requirements Tab)**
![File Structure](https://github.com/cfreeman21/images/blob/master/KMSFix_SCCM_5.png)<p>
