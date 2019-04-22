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

    Add-Content $Logfile -value $logstring
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
$CurrentVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "CurrentVersionDB"

LogWrite "Current Version DB= $CurrentVersion"

#Database Instance
$DBInstance = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "DBInstance"

LogWrite "DB Instance Name= $DBInstance"

#Database Server
$DBServer = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "DBServer"

LogWrite "DB Server Name= $DBServer"

#Install Location
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "InstallLocation"

LogWrite "Install Location= $Install"

#Server Name
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "Server"

LogWrite "Server Name= $ServerName"

#Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "ServiceAccount"

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
$EndUserName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\LicenseRegistryInfo" -Name "EndUserName"

LogWrite "End User Name= $EndUserName"

$EndUserPhone = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\LicenseRegistryInfo" -Name "EndUserPhone"

LogWrite "End User Phone= $EndUserPhone"

$EndUserEmail = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\LicenseRegistryInfo" -Name "EndUserEmail"

LogWrite "End User Email= $EndUserEmail"


#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............License Keys.............."

#License Address
$LicenseAddress = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseAddress"

LogWrite "License Address= $LicenseAddress"

#License Key Dates
$LicenseKeyASMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyASMDate"

LogWrite "License Key A Date= $LicenseKeyASMDate"

$LicenseKeyCSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyCSMDate"

LogWrite "License Key C Date= $LicenseKeyCSMDate"

$LicenseKeyNSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyNSMDate"

LogWrite "License Key N Date= $LicenseKeyNSMDate"

$LicenseKeyOSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyOSMDate"

LogWrite "License Key O Date= $LicenseKeyOSMDate"

$LicenseKeyPSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyPSMDate"

LogWrite "License Key P Date= $LicenseKeyPSMDate"

$LicenseKeyTSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyTSMDate"

LogWrite "License Key T Date= $LicenseKeyTSMDate"

#License Keys
$LicenseKeyA = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyA"

LogWrite "License Key A= $LicenseKeyA"

$LicenseKeyC = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyC"

LogWrite "License Key C= $LicenseKeyC"

$LicenseKeyD = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyD"

LogWrite "License Key D= $LicenseKeyD"

$LicenseKeyE = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyE"

LogWrite "License Key E= $LicenseKeyE"

$LicenseKeyG = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyG"

LogWrite "License Key G= $LicenseKeyG"

$LicenseKeyM = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyM"

LogWrite "License Key M= $LicenseKeyM"

$LicenseKeyN = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyN"

LogWrite "License Key N= $LicenseKeyN"

$LicenseKeyO = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyO"

LogWrite "License Key O= $LicenseKeyO"

$LicenseKeyP = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyP"

LogWrite "License Key P= $LicenseKeyP"

$LicenseKeyT = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange\Profile 1" -Name "LicenseKeyT"

LogWrite "License Key T= $LicenseKeyT"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............Tips.............."

LogWrite "
Project
On Premise
O365
Hybrid 365
Migration - Both on Premise and 365

Contact info
DidITBetter Support
813-977-5739 

Credentials
Sync Service Account

Account used for Permissions


Current Situation


Current Issues


Settings Sync Intervals

Add2Exchange GAL Manager � 

Add2Exchange Rel Manager � making and removing relationships - run at least once a day

Sync for Calendars � specified sleep time, between syncs

Sync for Contacts � specified sleep time between syncs 

Sync for Tasks � specified sleep time between syncs 

Current Results

GAL:
RELMAN:
Add2Exchange Calendar Synchronization (Add2Agent):
Add2Exchange Contacts Synchronization (Add2Agent): 
Add2Exchange Tasks Synchronization (Add2Agent): 
Add2Exchange Posts Synchronization (Add2Agent): 
Add2Exchange Notes Synchronization (Add2Agent): 


Onboarding
Run Diditbetter_Support_Menu.ps1 as PowerShell file from desktop 

FIRST: If using Relationship Group Manager, add the users to the distribution list and then immediately give permissions.

To give permissions to a single source or destination folder not in a distribution list, select that option and specify by email address.

Select which Exchange environment you are in

If you are on Premise, Select "not on Exchange server" for permissions to an On-premise Exchange server and enter name (unless you are on it) 
 - Press continue for the remote PowerShell warning
If on premise and not on Exchange Server, enter the netbios name of an Exchange server. [This may not run through your load balancer for Kerberos authentication or if PS Remote is not enabled]  If it fails, you can copy this file to the Exchange Server and run as your Organizational Administrator.
File Default Location:  C:\Program Files (x86)\OpenDoor Software�\Add2Exchange\Setup\PermissionsOnPremOrO365Combined.ps1

Log in as Organization or Global admin or Exchange Organizational Administrator.

Tip for Automatic Permissions
Consider giving your Sync Service account this desired membership and you can set up Timed Automatic Permissions and give permissions with this account automatically

Global Admin or Exchange Organizational Administrator login Credentials:


When asked for service account, use your sync service account, default to ZADD2EXCHANGE

Select how to give permissions
Select option 0 to give to all, or option 3 to apply to only members of a distribution list 

Distribution list names:
Do again until all are done

Again, consider using the automatic scheduled permissions option from Diditbetter_Support_Menu

Off boarding

To off board users, remove the users from the attached Distribution list and then don�t hide, delete or detach or litigate or remove the Outlook license if 365 until at least 24 hours later.  This gives the program enough time to detect the user has been removed from the distribution list (Relman runs successfully) and then removes the data and the relationship automatically.

 
If done correctly, until at least 24 hours after, so the relationship can be marked and then auto dereplicate.  If the timing didn�t work, and the relationship for the �dead� account is in the trash and RED, you can select from Global Options to �Empty the Trash� and it close the Console and when prompted, you can select to �Not Continue where you left off� which will start at the beginning and remove all relationships in the trash, and force delete the relationships in the trash. Let it run a complete cycle. If you open the Add2Exchange Console before the full sync, you will have to reselect and specify to �Empty the Trash� again.


Tips � If Outlook is installed for the profile, you must leave machine logged in (and locked if desired) as the sync service account user
Should be set to auto login.  This allows patch management to work correctly and continue syncing after a reboot.  For MS updates to be easily managed you can set this in the DidITBetter Support Menu.ps1 and the pw is encrypted. If the pw changes, run it again.  This auto login is primarily to recover after MS updates and a reboot.  Outlook does not need to be open for proper operation.

If using the automatic timed permissions option, use an account which password does not change, or give the service account proper permissions, since its password is set to never expire.

If in Office 365, consider making the service account a global admin.  If On Premise Exchange, consider making the service account an Exchange Organizational Administrator and set up Automatic Timed Permissions from the PowerShell �DidITBetter Support Menu.ps1� 


YTD


Some required and recommended finishing touches: 

If Contact Syncing (especially GAL) 

There is a default add in in Outlook add called the �Outlook Social Connector� which can cause the items in the destination to be changed.  
The effect is it is updating many of these contacts in every destination as evidenced by long relationship sync times.  A Ping pong effect occurs, where we sync to destination, the destination items change and then we change them back.  This is especially true if we make GAL entries as SMTP email address for Mobile utility. 
So for proper sync and put this continuing issue to rest, we immediately need to disable some or all parts of the Outlook Social Connector as group policy, and at least the option: �Blocking Global Address List synchronization with local contacts�. See MS link below 

After this policy is applied, it will require user logouts or reboots and back on of the machine or a forced group policy update, and closing and opening of Outlook.  After this, and two full syncs, we will be able to tell if other outside �off the domain� Outlook machines on an account, due to the longer times for the problem user.

Tip � if there are any Outlook clients on machines not logging into the domain (at home, for example) this will not get set by group policy, so should be disabled manually.  

It will also not be set if the Office 365 accounts are E1, and F1 accounts cannot be synced to as these are only online.  In order for Group Policy to be respected by the domain, they have to be Business or E3 accounts or better.  
Disabling can be done in the Add ins interface in Outlook, but often updates to MSO reenable those, so in the DidITBetter Support Menu.os1 (on the desktop) there is included the PowerShell script option to disable permanently by renaming the problem addins on machine � so that non domain type machine does not have the Social Connector enabled.  


How to Disable Outlook Social Connector by Group Policy
Our version: 
How to disable Outlook Social Connector by Group Policy or GP Edit
http://support.diditbetter.com/disable-group-policy.aspx

And Official Microsoft version: 
Manage the Outlook Social Connector with Group Policy
https://support.microsoft.com/en-us/help/2020103/how-to-manage-the-outlook-social-connector-by-using-group-policy


After this policy is applied it will require user logouts or reboots and back on of the machine, or a forced group policy update, and closing and opening of outlook.  After this, and two full syncs, we will be able to tell if other outside �off the domain� Outlook machines on an account, due to the longer times for that user.

How to Disable Outlook Social Connector by PowerShell on that machine

Copy the file from this location and run on the user�s machine
You may have to run as administrator and perhaps set this:
Set-executionpolicy -scope localmachine -executionpolicy unrestricted
This is the file, from the Replication server default location

C:\Program Files (x86)\OpenDoor Software�\Add2Exchange\Setup\Remove_Outlook_Add_ins.ps1

This PowerShell disables and renames the outlook addins so they are not accidentally reenabled.


Antivirus Exclusions from active file scanning
(These are default locations) 
C:\Program Files\OpenDoor Software�\
-	note the �Circle r� in the directory name. Copy and paste to antivirus program
C:\Program Files\Microsoft SQL Server\
C:\Zlibrary 



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