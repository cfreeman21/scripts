2 Lines to edit for the script is $LogDir:

$LogDir = "" # Example C:\CompanyXYZ\logs\WMF51
$Company = "" ## Company name this is used when creating registry key

WMF51-FullInstall.ps1 can be ran manually (Elevated CMD Prompt) or through Configuration Manager (SCCM)

# Manual Install:<h1><p>
(Open CMD as Administrator) Run [WMF51-FullInstall.ps1](https://github.com/cfreeman21/scripts/blob/master/WMF51/WMF51-FullInstall.cmd)<p>

# Configuration Manager (SCCM) Install:<h1><p>

Directory Structure:
![Directory Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/directory_structure.png)<p>

**Application Properties: (General Tab)**
![File Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/images/WMF51_SCCM_1.png)<p>
**Application Properties: (Programs Tab)**
![File Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/images/WMF51_SCCM_2.png)<p>
**Application Properties: (Detection Tab / Detection Rule)**
![File Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/images/WMF51_SCCM_3.png)<p>
CODE:
````if (($PSVersionTable.PSVersion | Select-Object -ExpandProperty Major) -eq 5 -and ($PSVersionTable.PSVersion | Select-Object -ExpandProperty Minor) -eq 1)
{
    Write-Host "Installed"
}
**Application Properties: (User Experience Tab)**
![File Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/images/WMF51_SCCM_4.png)<p>
**Application Properties: (Requirements Tab)**
![File Structure](https://github.com/cfreeman21/scripts/blob/master/WMF51/images/WMF51_SCCM_5.png)<p>
		Windows 7 SP1
		Windows 8.1
		Windows Server 2008 R2
		Windows Server 2012
		Windows Server 2012 R2
