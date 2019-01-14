<#	
	.NOTES
	===========================================================================
	 Created on:   	01/10/2019
	 Created by:   	Cory Freeman
	 Filename:     	KMS_Error_Fix.ps1
	===========================================================================
	.DESCRIPTION
		This File fixes issue in Enterprise Environment for KMS Activation "Windows not Genuine" The Script follows the process requested by Microsoft to resolve this issue.
#>
$LogDir = "" # Example C:\CompanyXYZ\logs\KMSFix
#Check if the folder Exist if not then create it
IF (!(Test-Path $LogDir))
{
	New-Item -Path $LogDir -ItemType Directory
}
ELSE
{
	Write-Host "Directory already exist"
}
#Log Files
$ScriptOutput = "$LogDir\ScriptOutput.log"
$SuccessLog = "$LogDir\Successful.log"
$ErrorLog = "$LogDir\Error.log"
$ScriptRunDate = Get-Date -Uformat %m-%d-%Y

#Function for CMTrace Log File Format
function Write-log
{
	
	[CmdletBinding()]
	Param (
		[parameter(Mandatory = $true)]
		[String]$Path,
		[parameter(Mandatory = $true)]
		[String]$Message,
		[parameter(Mandatory = $true)]
		[String]$Component,
		[Parameter(Mandatory = $true)]
		[ValidateSet("Info", "Warning", "Error")]
		[String]$Type
	)
	
	switch ($Type)
	{
		"Info" { [int]$Type = 1 }
		"Warning" { [int]$Type = 2 }
		"Error" { [int]$Type = 3 }
	}
	
	# Create a log entry
	$Content = "<![LOG[$Message]LOG]!>" +`
	"<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " +`
	"date=`"$(Get-Date -Format "M-d-yyyy")`" " +`
	"component=`"$Component`" " +`
	"context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
	"type=`"$Type`" " +`
	"thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
	"file=`"`">"
	
	# Write the line to the log file
	Add-Content -Path $Path -Value $Content
}

Write-log -Path $ScriptOutput -Message "----------Windows 7 KMS Fix Log Start----------" -Component "Script" -Type Info
Write-log -Path $ScriptOutput -Message "Script Run Date: $ScriptRunDate" -Component "Script" -Type Info

function RemoveUpdate
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$KB
	)
	
	$HotfixDetection = get-hotfix | Where-Object { $_.HotFixID -match "$KB" }
	IF ($HotfixDetection -ne $Null)
	{
		# Command line for removing the update
		$RemovalCommand = "$env:systemroot\System32\wusa.exe /uninstall /kb:$KB /quiet /norestart";
		Write-Host ("Removing update with command: " + $RemovalCommand);
		Write-log -Path $ScriptOutput -Message ("Removing update with command: " + $RemovalCommand) -Component "Script" -Type Info;
		
		# Invoke the command we built above
		Invoke-Expression -Command $RemovalCommand;
		
		# Wait for wusa.exe to finish and exit (wusa.exe actually leverages
		# TrustedInstaller.exe, so you won't see much activity within the wusa process)
		while (@(Get-Process wusa -ErrorAction SilentlyContinue).Count -ne 0)
		{
			Start-Sleep 1
			Write-Host "Waiting for update removal to finish ..."
			Write-log -Path $ScriptOutput -Message "Waiting for update removal to finish ..." -Component "Script" -Type Info;
		}
	}
	ELSE
	{
		Write-Host "KB$KB is not installed..."
		Write-log -Path $ScriptOutput -Message "KB$KB is not installed..." -Component "Script" -Type Info;
	}
	
}

function ServiceAction
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$ServiceName,
		[string]$Action
	)
	if (Get-Service $ServiceName -ErrorAction SilentlyContinue)
	{
		#Condition if user wants to stop a service
		if ($Action -eq 'Stop')
		{
			if ((Get-Service -Name $ServiceName).Status -eq 'Running')
			{
				Write-Host "$ServiceName is running, preparing to stop..."
				Write-log -Path $ScriptOutput -Message "$ServiceName is running, preparing to stop..." -Component "Script" -Type Info
				Get-Service -Name $ServiceName | Stop-Service -ErrorAction SilentlyContinue
				Start-Sleep 20
				$ServiceStatus = (Get-Service -Name $ServiceName).Status
				Write-Host "$ServiceName is $ServiceStatus"
				Write-log -Path $ScriptOutput -Message "$ServiceName is $ServiceStatus" -Component "Script" -Type Info
			}
			elseif ((Get-Service -Name $ServiceName).Status -eq 'Stopped')
			{
				Write-Host "$ServiceName already stopped!"
				Write-log -Path $ScriptOutput -Message "$ServiceName already stopped!" -Component "Script" -Type Info
			}
			else
			{
				Write-Host "$ServiceName is $ServiceStatus"
				Write-log -Path $ScriptOutput -Message "$ServiceName is $ServiceStatus" -Component "Script" -Type Info
			}
		}
		#Condition if user wants to start a service
		ELSEIF ($Action -eq 'Start')
		{
			if ((Get-Service -Name $ServiceName).Status -eq 'Running')
			{
				Write-Host $ServiceName "already running!"
				Write-log -Path $ScriptOutput -Message "$ServiceName already running!" -Component "Script" -Type Info
			}
			elseif ((Get-Service -Name $ServiceName).Status -eq 'Stopped')
			{
				Write-Host $ServiceName "is stopped, preparing to start..."
				Write-log -Path $ScriptOutput -Message "$ServiceName is stopped, preparing to start..." -Component "Script" -Type Info
				Get-Service -Name $ServiceName | Start-Service -ErrorAction SilentlyContinue
				Start-Sleep 20
				$ServiceStatus = (Get-Service -Name $ServiceName).Status
				Write-Host "$ServiceName is $ServiceStatus"
				Write-log -Path $ScriptOutput -Message "$ServiceName is $ServiceStatus" -Component "Script" -Type Info
			}
			else
			{
				Write-Host "$ServiceName is $ServiceStatus"
				Write-log -Path $ScriptOutput -Message "$ServiceName is $ServiceStatus" -Component "Script" -Type Info
			}
		}
	}
	else
	{
		Write-Host "$ServiceName not found"
		Write-log -Path $ScriptOutput -Message "$ServiceName not found" -Component "Script" -Type Info
	}
}

function removefiles
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$File
	)
	
	if (Test-Path $File)
	{
		Remove-Item -Path $File -Force
		$FileStatus = Test-Path $File
		if ($FileStatus -eq $False)
		{
			Write-host -foregroundcolor Green $File " Removed"
			Write-log -Path $ScriptOutput -Message "$File Removed" -Component "Script" -Type Info
		}
		ELSE
		{
			Write-host -foregroundcolor Red $File " Unable to Remove"
			Write-log -Path $ScriptOutput -Message "$File Unable to Remove" -Component "Script" -Type Info
		}
	}
	ELSE
	{
		Write-host -foregroundcolor Red $File " does not exist"
		Write-log -Path $ScriptOutput -Message "$File does not exist" -Component "Script" -Type Info
	}
}

#Remove KB971033 Breaks KMS in Enterprise Environment
RemoveUpdate -KB 971033

#Stop Software Protection Service
ServiceAction -ServiceName sppsvc -Action Stop
#Stop SPP Notification Service
ServiceAction -ServiceName sppuinotify -Action Stop
#Set SPP Notification Service to Disabled
Set-Service sppuinotify -StartupType Disabled

#Files to be deleted for KMS Fix per Microsoft Support
$filelist = @(
	("$env:systemroot\System32\7B296FB0-376B-497e-B012-9C450E1B7327-5P-0.C7483456-A289-439d-8115-601632D005A0"),
	("$env:systemroot\System32\7B296FB0-376B-497e-B012-9C450E1B7327-5P-1.C7483456-A289-439d-8115-601632D005A0"),
	("$env:systemroot\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\tokens.dat"),
	("$env:systemroot\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\cache\cache.dat")
)
#Delete the Files
foreach ($File in $filelist)
{
	removefiles -File $File
}

#Start Software Protection Service
ServiceAction -ServiceName sppsvc -Action Start

#Set Windows 7 Enterprise KMS Key
cscript.exe $env:systemroot\system32\slmgr.vbs /ipk 33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
#Register with KMS Server to Activate
cscript.exe $env:systemroot\system32\slmgr.vbs /ato

#Set SPP Notification Service to Manual
Set-Service sppuinotify -StartupType Manual

#Verify Windows is Genuine
#Windows Activation
$LicenseInformation = cscript "$env:SystemRoot\system32\slmgr.vbs" -dli
#if ((cscript "$env:SystemRoot\system32\slmgr.vbs" -dli) -Contains "License Status: Licensed")
if (($LicenseInformation) -Contains "License Status: Licensed")
{
	Write-Host -foregroundcolor Green "Activation Sucessful"
	Write-log -Path $ScriptOutput -Message "Activation Sucessful" -Component "Script" -Type Info
	Write-log -Path $SuccessLog -Message "Script Run Date: $ScriptRunDate" -Component "Script" -Type Info
	Write-log -Path $SuccessLog -Message "Activation Sucessful" -Component "Script" -Type Info
	Write-log -Path $SuccessLog -Message "$LicenseInformation" -Component "Script" -Type Info
	Write-log -Path $ScriptOutput -Message "$LicenseInformation" -Component "Script" -Type Info
	Write-log -Path $ScriptOutput -Message "----------Windows 7 KMS Fix Log End----------" -Component "Script" -Type Info
	Exit 0
}
else
{
	Write-Host -foregroundcolor Red "Activation Error"
	Write-log -Path $ScriptOutput -Message "Activation Error" -Component "Script" -Type Info
	Write-log -Path $ErrorLog -Message "Script Run Date: $ScriptRunDate" -Component "Script" -Type Info
	Write-log -Path $ErrorLog -Message "Activation Error" -Component "Script" -Type Error
	Write-log -Path $ErrorLog -Message "$LicenseInformation" -Component "Script" -Type Info
	Write-log -Path $ScriptOutput -Message "$LicenseInformation" -Component "Script" -Type Info
	Write-log -Path $ScriptOutput -Message "----------Windows 7 KMS Fix Log End----------" -Component "Script" -Type Info
	Exit 1
}
