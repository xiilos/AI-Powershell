if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Logging
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}

$Logfile = "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Setup_Details.log"
Function LogWrite {
    Param ([string]$logstring)

    Add-content $Logfile -value $logstring
}


#Clear The Log
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Setup_Details.log"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Log File Exists..."
    $confirmation = Read-Host "Would you like to clear the Log? [Y/N]"
if ($confirmation -eq 'N') {
    Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
    Clear-Content "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Setup_Details.log"
}
}


LogWrite "..............Add2Exchange Details.............."

#Time and Date
$Start_Time = Get-Date
LogWrite "Date/Time= $Start_Time"

#Windows Version
$WindowsVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption

LogWrite "Windows Version= $WindowsVersion"

#Domain Name
$Domain = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "Domain"

LogWrite "Domain Name= $Domain"

#PowerShell Version
$PShellVersion = $PSVersionTable.PSVersion

LogWrite "PowerShell Version= $PShellVersion"

#Database Version
$CurrentVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "CurrentVersionDB"

LogWrite "Current Version DB= $CurrentVersion"

#Database Instance
$DBInstance = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance"

LogWrite "DB Instance Name= $DBInstance"

#Database Server
$DBServer = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBServer"

LogWrite "DB Server Name= $DBServer"

#Install Location
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation"

LogWrite "Install Location= $Install"

#Server Name
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server"

LogWrite "Server Name= $ServerName"

#Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount"

LogWrite "Service Account Name= $ServiceAccount"

#ServiceAccount Password
#$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerPass.txt" | convertto-securestring
$Password = Read-Host "What is the Service Account password?"
LogWrite "Service Account Password= $Password"

#Connection Method
if ($DBServer -like $ServerName) {
    LogWrite "Connection Method= Outlook"
}

Else {
    LogWrite "Connection Method= CDO or Remote SQL"
}
#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............End User Information.............."

#EndUser Information
$EndUserName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserName"

LogWrite "End User Name= $EndUserName"

$EndUserPhone = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserPhone"

LogWrite "End User Phone= $EndUserPhone"

$EndUserEmail = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserEmail"

LogWrite "End User Email= $EndUserEmail"


#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............License Keys.............."

#License Address
$LicenseAddress = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseAddress"

LogWrite "License Address= $LicenseAddress"

#License Key Dates
$LicenseKeyASMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyASMDate"

LogWrite "License Key A Date= $LicenseKeyASMDate"

$LicenseKeyCSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCSMDate"

LogWrite "License Key C Date= $LicenseKeyCSMDate"

$LicenseKeyNSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNSMDate"

LogWrite "License Key N Date= $LicenseKeyNSMDate"

$LicenseKeyOSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOSMDate"

LogWrite "License Key O Date= $LicenseKeyOSMDate"

$LicenseKeyPSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPSMDate"

LogWrite "License Key P Date= $LicenseKeyPSMDate"

$LicenseKeyTSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTSMDate"

LogWrite "License Key T Date= $LicenseKeyTSMDate"

#License Keys
$LicenseKeyA = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyA"

LogWrite "License Key A= $LicenseKeyA"

$LicenseKeyC = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyC"

LogWrite "License Key C= $LicenseKeyC"

$LicenseKeyD = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyD"

LogWrite "License Key D= $LicenseKeyD"

$LicenseKeyE = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyE"

LogWrite "License Key E= $LicenseKeyE"

$LicenseKeyG = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyG"

LogWrite "License Key G= $LicenseKeyG"

$LicenseKeyM = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyM"

LogWrite "License Key M= $LicenseKeyM"

$LicenseKeyN = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyN"

LogWrite "License Key N= $LicenseKeyN"

$LicenseKeyO = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyO"

LogWrite "License Key O= $LicenseKeyO"

$LicenseKeyP = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyP"

LogWrite "License Key P= $LicenseKeyP"

$LicenseKeyT = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyT"

LogWrite "License Key T= $LicenseKeyT"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............Tips.............."

LogWrite "
Onboarding
If using Relationship Group Manager, Add the users to the distribution list and then immediately
give permissions

Run Diditbetter_Support_Menu as Powershell file from desktop 

Select which Exchange environment you are in

Select *not on exchange server* for permissions to an On-prem Exchange server name  
Press continue for the remote powershell warning
Enter netbios name of an exchange server, this will will not run through most load balancers for kerberos authentication
Log in as organization or tenant admin

Global Admin or Exchange Organizational Administrator login Credentials:


When asked for service account
use your sync service account, default to ZADD2EXCHANGE

Select how to give permissions
Select option 0 to give to all, or option 3 to apply to only members of a distribution list 

Distribution list names:
do again until all are done

Consider using the automatic scheduled permissions option from Diditbetter_Support_Menu



Tips – must leave machine logged in (and locked if desired)  as the sync service account user
Must set to autologin for MS updates to be easily managed

If using automatic permissions, use an account which password does not change.
If in Office 365, Consider making the service account a global admin.





Current Situation

Current Issues

Current Results
Add2Exchange GAL Manager – 

Add2Exchange Rel Manager – making and removing rels

Add2Exchange Calendar Synchronization (Add2Agent)  

Add2Exchange Contacts Synchronization (Add2Agent) 

Add2Exchange Tasks Synchronization (Add2Agent) 

Sync for GAL 


Sync Intervals

Relman 
run at least once a day

Sync for Calendars – specified sleep time, between syncs

Sync for Contacts – specified sleep time between syncs when others are not syncing

Sync for Tasks – specified sleep time between syncs when others are not syncing



YTD

Some finishing touches after we fix your issues

If Contact Syncing (especially GAL) 
There is a default add in in Outlook add called the Outlook Social Connector which can cause the items in the destination to be changed.  
The effect is it is updating many of these contacts in every destination as evidenced by this relationship
So for proper sync and put this continuing issue to rest,  we immediately need to disable some or all parts of the outlook social connector as group policy, at least the option for GAL Sync - the best way.
Tip – if there are any outlook clients on machines not logging into the domain (by machine) this will not get set, so should be disabled manually.  It can be done in the Add ins interface in Outlook, but often updates to MSO reenable those, so I will include the powershell script to disable permanently on that non logged into the domain type machine if there are any.
How to Disable Outlook Social Connector
Our version: http://support.diditbetter.com/disable-group-policy.aspx
And MS version: Manage the Outlook Social Connector with Group Policy
After this it will require user logouts or reboots and back on of the machine 
or a forced group policy update, and closing and opening of outlook.  After this, and a full sync, we will be able to tell if other outside “off the domain” Outlook machines on an account, due to the longer times for that user.

Antivirus Exclusions from active file scanning
Folder and subfolders
C:\Program Files\OpenDoor Software®\
C:\Program Files\Microsoft SQL Server\
C:\Zlibrary or wherever you put our setup files 
or
C:\Program Files\My Documents – wherever that lives for this sync user



History
"
Start-Sleep -Seconds 2
#Shortcut
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Support.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Setup_Details.log"
$Shortcut.Save()

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting