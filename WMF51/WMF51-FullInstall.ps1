<#	
	.NOTES
	===========================================================================
	 Created on:   	3/7/2017
	 Created by:   	Cory Freeman
	 Filename:     	WMF51-FullInstall.ps1
	===========================================================================
	.DESCRIPTION
		This script is used to Check All PreReqs for WMF 5.1 for Supported Operating Systems
	.OPERATINGSYSTEMS
		Windows 7 SP1
		Windows 8.1
		Windows Server 2008 R2
		Windows Server 2012
		Windows Server 2012 R2
	.PREREQS
		Latest service pack must be installed.
		WMF 3.0 must not be installed. Installing WMF 5.1 over WMF 3.0 will result in the loss of the PSModulePath, which can cause other applications to fail. Before installing WMF 5.1, you must either un-install WMF 3.0, or save the PSModulePath and then restore it manually after WMF 5.1 installation is complete.
		WMF 5.1 requires .NET Framework 4.5.2 You can install Microsoft .NET Framework 4.5.2 by following the instructions at the download location.
	.REFERENCELINK
		https://msdn.microsoft.com/en-us/powershell/wmf/5.1/install-configure
	
#>
[CmdletBinding()]
param()
## Global Variables
$LogDir = "" # Example C:\CompanyXYZ\logs\WMF51
#Check if the folder Exist if not then create it
IF (!(Test-Path $LogDir))
{
	New-Item -Path $LogDir -ItemType Directory
}
ELSE
{
	Write-Host "Directory already exist"
}
Start-Transcript -path $LogDir\wmf51_transcript.txt -append
$Log = "$LogDir\WMF51.log" ## Log Location from Script Output
Write-Host $Log
$ScriptRunDate = Get-Date -Uformat %m-%d-%Y
Write-Host $ScriptRunDate
$ScriptDir = split-path -Path $MyInvocation.MyCommand.Definition -Parent
Write-Host $ScriptDir
$Company = "" ## Company name this is used when creating registry key
$registryPath = "HKLM:\Software\$Company\WMF51" ## Registry Path to store the WMF Install Status
$ScriptRunBy = whoami
############################### DO NOT EDIT BELOW ########################################################

Function Log-ScriptEvent
{
	
	#########################################################################################################################################
	# This Funciton Came from https://gallery.technet.microsoft.com/scriptcenter/Log-ScriptEvent-Function-ea238b85 All Credit to the Author #
	#########################################################################################################################################
	
	#Define and validate parameters
[CmdletBinding()]
Param(
      #Path to the log file
      [parameter(Mandatory=$True)]
      [String]$NewLog,

      #The information to log
      [parameter(Mandatory=$True)]
      [String]$Value,

      #The source of the error
      [parameter(Mandatory=$True)]
      [String]$Component,

      #The severity (1 - Information, 2- Warning, 3 - Error)
      [parameter(Mandatory=$True)]
      [ValidateRange(1,3)]
      [Single]$Severity
      )


#Obtain UTC offset
$DateTime = New-Object -ComObject WbemScripting.SWbemDateTime 
$DateTime.SetVarDate($(Get-Date))
$UtcValue = $DateTime.Value
$UtcOffset = $UtcValue.Substring(21, $UtcValue.Length - 21)


#Create the line to be logged
$LogLine =  "<![LOG[$Value]LOG]!>" +`
            "<time=`"$(Get-Date -Format HH:mm:ss.fff)$($UtcOffset)`" " +`
            "date=`"$(Get-Date -Format M-d-yyyy)`" " +`
            "component=`"$Component`" " +`
            "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
            "type=`"$Severity`" " +`
            "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
            "file=`"`">"

#Write the line to the passed log file
Add-Content -Path $NewLog -Value $LogLine

}

##########################################################################################################

##########################################################################################################
##Start Logging Information
		Log-ScriptEvent $Log "----------WMF 5.1 Full Install Log Start----------" -Component "WMF 5.1" -Severity 1
Log-ScriptEvent $Log "Script Run Date: $ScriptRunDate" -Component "WMF 5.1" -Severity 1

#WMF 5.1 Supported Operating System Check
Function Get-OperatingSystem
{
	$Version = (Get-WMIObject -Class Win32_OperatingSystem).version
	$ProductType = (Get-WMIObject -Class Win32_OperatingSystem).ProductType
	$OSArchitecture = (Get-WMIObject -Class Win32_OperatingSystem).OSArchitecture
	
	IF ($Version -like "6.1*" -AND $ProductType -eq "1" -AND $OSArchitecture -ne "64-bit")
	{
		$script:OSVersion = "Win7x86"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win7-KB3191566-x86.msu"
		$script:HotFixID = "KB3191566"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.1*" -AND $ProductType -eq "1" -AND $OSArchitecture -eq "64-bit")
	{
		$script:OSVersion = "Win7x64"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win7AndW2K8R2-KB3191566-x64.msu"
		$script:HotFixID = "KB3191566"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.3*" -AND $ProductType -eq "1" -AND $OSArchitecture -ne "64-bit")
	{
		$script:OSVersion = "Win81x86"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win8.1-KB3191564-x86.msu"
		$script:HotFixID = "KB3191564"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.3*" -AND $ProductType -eq "1" -AND $OSArchitecture -eq "64-bit")
	{
		$script:OSVersion = "Win81x64"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win8.1AndW2K12R2-KB3191564-x64.msu"
		$script:HotFixID = "KB3191564"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.3*" -AND $ProductType -ne "1")
	{
		$script:OSVersion = "Win2008R2x64"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win7AndW2K8R2-KB3191566-x64.msu"
		$script:HotFixID = "KB3191566"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.2*" -AND $ProductType -ne "1")
	{
		$script:OSVersion = "Win2012x64"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "W2K12-KB3191565-x64.msu"
		$script:HotFixID = "KB3191565"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($Version -like "6.3*" -AND $ProductType -ne "1")
	{
		$script:OSVersion = "Win2012R2x64"
		$script:OSVersionSupported = "YES"
		$script:WMF51MSU = "Win8.1AndW2K12R2-KB3191564-x64.msu"
		$script:HotFixID = "KB3191564"
		Log-ScriptEvent $Log "Operating System: $script:OSVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSE
	{
		$script:OSVersionSupported = "NO"
		Log-ScriptEvent $Log "The Script is Exiting now, Unsupported Operating System" -Component "WMF 5.1" -Severity 3
		EXIT 0
	}
}

# Check if .Net 4.5 or above is installed for all Operating Supported Operating Systems for WMF 5.1 Install (No Windows 10 .Net Checks)
# Release Numbers from https://msdn.microsoft.com/en-us/library/hh925568%28v=vs.110%29.aspx
Function Get-NetFrameworkVersion
{
	$dotNet4Registry = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release).Release
	
	IF ($dotNet4Registry -eq 378389)
	{
		$script:NetVersion = "4.5"
		$script:NetVersionSupported = "NO"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 378675)
	{
		$script:NetVersion = "4.5.1"
		$script:NetVersionSupported = "NO"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 378758)
	{
		$script:NetVersion = "4.5.1"
		$script:NetVersionSupported = "NO"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 379893)
	{
		$script:NetVersion = "4.5.2"
		$script:NetVersionSupported = "YES"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 393297)
	{
		$script:NetVersion = "4.6"
		$script:NetVersionSupported = "YES"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 394271)
	{
		$script:NetVersion = "4.6.1"
		$script:NetVersionSupported = "YES"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSEIF ($dotNet4Registry -eq 394806)
	{
		$script:NetVersion = "4.6.2"
		$script:NetVersionSupported = "YES"
		Log-ScriptEvent $Log ".Net Framework Version: $script:NetVersion" -Component "WMF 5.1" -Severity 1
	}
	ELSE
	{
		$script:NetVersionSupported = "NO"
		Log-ScriptEvent $Log "The Script is Exiting now, Unsupported .Net Framework" -Component "WMF 5.1" -Severity 3
		EXIT 1
	}
}

Function Test-OSVersion
{
	if ($OSVersionSupported -eq "YES")
	{
		return $true
	}
	return $false
}

Function Test-NetVersion
{
	if ($NetVersionSupported -eq "YES")
	{
		return $true
	}
	return $false
}

Function Test-PSVersion
{
	$REQPSVersion = "5.1"
	$PSMajor = $PSVersionTable.PSVersion.Major
	$PSMinor = $PSVersionTable.PSVersion.Minor
	$PSVersion = "$PSMajor" + "." + "$PSMinor"
	if ($PSVersion -eq $REQPSVersion)
	{
		return $true
	}
	return $false
}

#Check if WMF 3.0 is installed only on supported OS Versions
Function Test-WMF3Version
{
	IF (((($OSVersion -eq "Win7x86") -OR ($OSVersion -eq "Win7x64") -OR ($OSVersion -eq "Win2008R2x64") -AND (Get-Hotfix "KB2506143"))))
	{
		return $true
	}
	else
	{
		return $false
	}
}

Get-OperatingSystem
Get-NetFrameworkVersion

$OSVER = Test-OSVersion
Write-Host "OSVersion: $OSVER" #True means supported OS, False means not supported OS
Log-ScriptEvent $Log "OSVersion: $OSVER" -Component "WMF 5.1" -Severity 1
$NETVER = Test-NetVersion
Write-Host "NETVersion: $NETVER" #True means supported .Net Version, False means not supported .Net Version
Log-ScriptEvent $Log "NETVersion: $NETVER" -Component "WMF 5.1" -Severity 1
$PSVER = Test-PSVersion
Write-Host "PSVersion: $PSVER" #True means WMF 5.1 alerady installed, False means older version of Powershell exist
Log-ScriptEvent $Log "PSVersion: $PSVER" -Component "WMF 5.1" -Severity 1
$WMF3VER = Test-WMF3Version
Write-Host "WMF3Version: $WMF3VER" #True means WMF3 is installed, False means WMF3 is not installed
Log-ScriptEvent $Log "WMF3Version: $WMF3VER" -Component "WMF 5.1" -Severity 1

#Run Through all Pre-Req Check and Install whats needed
	IF ($PSVER -eq $true)
	{
		Log-ScriptEvent $Log "Powershell Version 5.1 is already installed" -Component "WMF 5.1" -Severity 1
		Write-Host "Powershell Version 5.1 is already installed"
		Log-ScriptEvent $Log "Do Nothing Else" -Component "WMF 5.1" -Severity 1
		EXIT 0
	}
	ELSEIF (((($PSVER -eq $false) -AND ($OSVER -eq $true) -AND ($NETVER -eq $true) -AND ($WMF3VER -eq $false))))
	{
		Log-ScriptEvent $Log "All Prerequisites have pass and ready for Install" -Component "WMF 5.1" -Severity 1
		Write-Host "All Prerequisites have pass and ready for Install"
		Write-Host "Installing WMF 5.1"
		Log-ScriptEvent $Log "Installing WMF 5.1" -Component "WMF 5.1" -Severity 1
		Log-ScriptEvent $Log "Installing $WMF51MSU for $OSVersion" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "$ScriptDir\$WMF51MSU /quiet /norestart /logfile:$LogDir\WMF51_INSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
	}
	ELSEIF (($PSVER -eq $false) -AND ($OSVER -eq $false))
	{
		Log-ScriptEvent $Log "Unsupported OS" -Component "WMF 5.1" -Severity 3
		Write-Host "Unsupported OS"
	}
	ELSEIF (((($PSVER -eq $false) -AND ($OSVER -eq $true) -AND ($NETVER -eq $false) -AND ($WMF3VER -eq $false))))
	{
		Log-ScriptEvent $Log "Unsupported .Net Framework" -Component "WMF 5.1" -Severity 2
		Log-ScriptEvent $Log ".Net 4.5.2 or Above is needed" -Component "WMF 5.1" -Severity 2
		Write-Host "Unsupported .Net Framework"
		Write-Host ".Net 4.5.2 or Above is needed"
		Write-Host "Installing .Net 4.5.2"
		Log-ScriptEvent $Log "Installing .Net 4.5.2" -Component "WMF 5.1" -Severity 1
		(Start-Process -FilePath '.\NDP452-KB2901907-x86-x64-AllOS-ENU.exe' -Argumentlist "/norestart /q /log $LogDir\NET_452_Install.log" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 10 Minutes to allow install to finish
		Start-Sleep -s 600
		Write-Host "Installing WMF 5.1"
		Log-ScriptEvent $Log "Installing $WMF51MSU for $OSVersion" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "$ScriptDir\$WMF51MSU /quiet /norestart /logfile:$LogDir\WMF51_INSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
	}
	ELSEIF (((($PSVER -eq $false) -AND ($OSVER -eq $true) -AND ($NETVER -eq $true) -AND ($WMF3VER -eq $true))))
	{
		Log-ScriptEvent $Log "WMF 3.0 is Installed and needs to be uninstalled" -Component "WMF 5.1" -Severity 2
		Write-Host "WMF 3.0 is Installed and needs to be uninstalled"
		Write-Host "Uninstalling WMF 3"
		Log-ScriptEvent $Log "Uninstalling WMF 3" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "/uninstall /KB:2506143 /quiet /norestart /logfile:$LogDir\WMF3_UNINSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
		Write-Host "Installing WMF 5.1"
		Log-ScriptEvent $Log "Installing $WMF51MSU for $OSVersion" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "$ScriptDir\$WMF51MSU /quiet /norestart /logfile:$LogDir\WMF51_INSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
	}
	ELSEIF (((($PSVER -eq $false) -AND ($OSVER -eq $true) -AND ($NETVER -eq $false) -AND ($WMF3VER -eq $true))))
	{
		Log-ScriptEvent $Log ".Net Needs to be updated and WMF 3.0 is Installed and needs to be uninstalled" -Component "WMF 5.1" -Severity 2
		Log-ScriptEvent $Log "Unsupported .Net Framework" -Component "WMF 5.1" -Severity 2
		Log-ScriptEvent $Log ".Net 4.5.2 or Above is needed" -Component "WMF 5.1" -Severity 2
		Write-Host "Unsupported .Net Framework"
		Write-Host ".Net 4.5.2 or Above is needed"
		Write-Host "Installing .Net 4.5.2"
		Log-ScriptEvent $Log "Installing .Net 4.5.2" -Component "WMF 5.1" -Severity 1
		(Start-Process -FilePath '.\NDP452-KB2901907-x86-x64-AllOS-ENU.exe' -Argumentlist "/norestart /q /log $LogDir\NET_452_Install.log" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 600
		Write-Host "WMF 3.0 is Installed and needs to be uninstalled"
		Write-Host "Uninstalling WMF 3"
		Log-ScriptEvent $Log "Uninstalling WMF 3" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "/uninstall /KB:2506143 /quiet /norestart /logfile:$LogDir\WMF3_UNINSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
		Write-Host "Installing WMF 5.1"
		Log-ScriptEvent $Log "Installing $WMF51MSU for $OSVersion" -Component "WMF 5.1" -Severity 1
		(Start-Process wusa.exe -ArgumentList "$ScriptDir\$WMF51MSU /quiet /norestart /logfile:$LogDir\WMF51_INSTALL_Update.evtx" -Wait -Passthru).ExitCode | Out-Null
		#Sleep for 5 Minutes to allow install to finish
		Start-Sleep -s 300
	}
	
	#Check if WMF 5.1 is Installed and report Success or Failure
	$VerifyWMF51 = Get-Hotfix $HotFixID
	IF ($VerifyWMF51 -ne $NULL)
	{
		Log-ScriptEvent $Log "WMF 5.1 Has successfully been installed" -Component "WMF 5.1" -Severity 1
		Log-ScriptEvent $Log "Reboot Needed" -Component "WMF 5.1" -Severity 1
		Log-ScriptEvent $Log "----------WMF 5.1 Full Install Log End----------" -Component "WMF 5.1" -Severity 1
		#Brand Registry for WMF Install Status
		IF (!(Test-Path $registryPath))
		{
		New-Item -Path $registryPath -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Install Date" -Value $ScriptRunDate -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Log Directory" -Value $LogDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran From" -Value $ScriptDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran By" -Value $ScriptRunBy -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Status" -Value "Success" -PropertyType STRING -Force | Out-Null
		}
		else
		{
		New-ItemProperty -Path $registryPath -Name "Install Date" -Value $ScriptRunDate -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Log Directory" -Value $LogDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran From" -Value $ScriptDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran By" -Value $ScriptRunBy -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Status" -Value "Success" -PropertyType STRING -Force | Out-Null
		}
		Write-Host "WMF 5.1 Has successfully been installed" -ForegroundColor Green
		EXIT 0
		Stop-Transcript
	}
	else
	{
		Log-ScriptEvent $Log "The Script has errored please review the logs $Log" -Component "WMF 5.1" -Severity 1
		Log-ScriptEvent $Log "Reboot Needed" -Component "WMF 5.1" -Severity 1
		Log-ScriptEvent $Log "----------WMF 5.1 Full Install Log End----------" -Component "WMF 5.1" -Severity 1
		IF (!(Test-Path $registryPath))
		{
		New-Item -Path $registryPath -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Install Date" -Value $ScriptRunDate -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Log Directory" -Value $LogDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran From" -Value $ScriptDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran By" -Value $ScriptRunBy -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Status" -Value "FAIL" -PropertyType STRING -Force | Out-Null		
		}
		else
		{		
		New-ItemProperty -Path $registryPath -Name "Install Date" -Value $ScriptRunDate -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Log Directory" -Value $LogDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran From" -Value $ScriptDir -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Script Ran By" -Value $ScriptRunBy -PropertyType STRING -Force | Out-Null
		New-ItemProperty -Path $registryPath -Name "Status" -Value "FAIL" -PropertyType STRING -Force | Out-Null
		}
		Write-Host "The Script has errored please review the logs $Log" -ForegroundColor Red
		EXIT 1
		Stop-Transcript
	}

