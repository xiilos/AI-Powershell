<#
        .SYNOPSIS
        Powershell script to include Add2Exchange setup details in .txt

        .DESCRIPTION
        Checks registry for A2E setup details and prints to .txt file
        Get licensing info
        Get install paths
        Get local account for Add2Exchange
        Get PS Version
        Get Windows Version
        Get DB Version


        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>


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

    Add-Content $Logfile -Value $logstring
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

#Date and Time
$Start_Time = Get-Date -ErrorAction SilentlyContinue
LogWrite "Date/Time= $Start_Time" -ErrorAction SilentlyContinue

#Windows Version
$WindowsVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption

LogWrite "Windows Version= $WindowsVersion" -ErrorAction SilentlyContinue

#Install Location
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue

LogWrite "Install Location= $Install" -ErrorAction SilentlyContinue

#Domain Name
$Domain = $env:UserDomain

LogWrite "Domain Name= $Domain" -ErrorAction SilentlyContinue

#Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount" -ErrorAction SilentlyContinue

LogWrite "Service Account Name= $ServiceAccount" -ErrorAction SilentlyContinue

#ServiceAccount Password
$Password = Read-Host "What is the Service Account password?" -ErrorAction SilentlyContinue

LogWrite "Service Account Password= $Password" -ErrorAction SilentlyContinue

#Local Account Name
$LocalAccount = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

LogWrite "Local Account Name= $LocalAccount" -ErrorAction SilentlyContinue

#Local Account Password

$LocalPassword = Read-Host "What is the Local Account password?" -ErrorAction SilentlyContinue

LogWrite "Local Account Password= $LocalPassword" -ErrorAction SilentlyContinue

#Computer Name
$CompName = $env:ComputerName

LogWrite "Computer Name= $CompName" -ErrorAction SilentlyContinue


#Logon Method
$Logon = Read-Host "Type In your Logon Method. Ex. Exchange 2016, or Office 365" -ErrorAction SilentlyContinue

LogWrite "Logon Method= $Logon" -ErrorAction SilentlyContinue

#Connection Method
if ($DBServer -like $ServerName) {
    LogWrite "Connection Method= Outlook"
}

Else {
    LogWrite "Connection Method= CDO or Remote SQL"
}

#License Address
$LicenseAddress = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseAddress" -ErrorAction SilentlyContinue

LogWrite "License Address= $LicenseAddress" -ErrorAction SilentlyContinue

#Database Server
$DBServer = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBServer" -ErrorAction SilentlyContinue

LogWrite "DB Server Name= $DBServer" -ErrorAction SilentlyContinue

#Server Name
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server" -ErrorAction SilentlyContinue

LogWrite "Server Name= $ServerName" -ErrorAction SilentlyContinue

#PowerShell Version
$PShellVersion = $PSVersionTable.PSVersion

LogWrite "PowerShell Version= $PShellVersion" -ErrorAction SilentlyContinue

#Database Version
$CurrentVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "CurrentVersionDB" -ErrorAction SilentlyContinue

LogWrite "Current Version DB= $CurrentVersion" -ErrorAction SilentlyContinue

#Database Instance
$DBInstance = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue

LogWrite "DB Instance Name= $DBInstance" -ErrorAction SilentlyContinue

#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............End User Information.............."

#EndUser Information
$EndUserName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserName" -ErrorAction SilentlyContinue

LogWrite "End User Name= $EndUserName" -ErrorAction SilentlyContinue

$EndUserPhone = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserPhone" -ErrorAction SilentlyContinue

LogWrite "End User Phone= $EndUserPhone" -ErrorAction SilentlyContinue

$EndUserEmail = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" -Name "EndUserEmail" -ErrorAction SilentlyContinue

LogWrite "End User Email= $EndUserEmail" -ErrorAction SilentlyContinue

#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............License Keys.............."

#License Keys & Dates
$LicenseKeyA = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyA" -ErrorAction SilentlyContinue

LogWrite "License Key A= $LicenseKeyA" -ErrorAction SilentlyContinue

$LicenseKeyAExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyAExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key A Date= $LicenseKeyAExpiry" -ErrorAction SilentlyContinue



$LicenseKeyC = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyC" -ErrorAction SilentlyContinue

LogWrite "License Key C= $LicenseKeyC" -ErrorAction SilentlyContinue

$LicenseKeyCExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key C Date= $LicenseKeyCExpiry" -ErrorAction SilentlyContinue



$LicenseKeyD = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyD" -ErrorAction SilentlyContinue

LogWrite "License Key D= $LicenseKeyD" -ErrorAction SilentlyContinue

$LicenseKeyDExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyDExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key D Date= $LicenseKeyDExpiry" -ErrorAction SilentlyContinue



$LicenseKeyE = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyE" -ErrorAction SilentlyContinue

LogWrite "License Key E= $LicenseKeyE" -ErrorAction SilentlyContinue

$LicenseKeyEExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyEExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key E Date= $LicenseKeyEExpiry" -ErrorAction SilentlyContinue



$LicenseKeyG = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyG" -ErrorAction SilentlyContinue

LogWrite "License Key G= $LicenseKeyG" -ErrorAction SilentlyContinue

$LicenseKeyGExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyGExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key G Date= $LicenseKeyGExpiry" -ErrorAction SilentlyContinue



$LicenseKeyM = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyM" -ErrorAction SilentlyContinue

LogWrite "License Key M= $LicenseKeyM" -ErrorAction SilentlyContinue

$LicenseKeyMExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyMExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key M Date= $LicenseKeyMExpiry" -ErrorAction SilentlyContinue



$LicenseKeyN = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyN" -ErrorAction SilentlyContinue

LogWrite "License Key N= $LicenseKeyN" -ErrorAction SilentlyContinue

$LicenseKeyNExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key N Date= $LicenseKeyNExpiry" -ErrorAction SilentlyContinue



$LicenseKeyO = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyO" -ErrorAction SilentlyContinue

LogWrite "License Key O= $LicenseKeyO" -ErrorAction SilentlyContinue

$LicenseKeyOExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key O Date= $LicenseKeyOExpiry" -ErrorAction SilentlyContinue



$LicenseKeyP = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyP" -ErrorAction SilentlyContinue

LogWrite "License Key P= $LicenseKeyP" -ErrorAction SilentlyContinue

$LicenseKeyPExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key P Date= $LicenseKeyPExpiry" -ErrorAction SilentlyContinue



$LicenseKeyT = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyT" -ErrorAction SilentlyContinue

LogWrite "License Key T= $LicenseKeyT" -ErrorAction SilentlyContinue

$LicenseKeyTExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTExpiry" -ErrorAction SilentlyContinue

LogWrite "License Key T Date= $LicenseKeyTExpiry" -ErrorAction SilentlyContinue

#----------------------------------------------------------------------------------------------------------------------------------------------------------------
LogWrite "..............Information and Notes.............."

LogWrite "

--------Onboarding--------

Run DidTtbetter Support Menu.ps1 as PowerShell file from desktop 

FIRST: If using Relationship Group Manager, add the users to the distribution list and then immediately give permissions.

To give permissions to a single source or destination folder not in a distribution list, select that option and specify by email address.

Select which Exchange environment you are in

If you are on Premise, Select *not on Exchange server* for permissions to an On-premise Exchange server and enter name (unless you are on it) 
 - Press continue for the remote PowerShell warning
If on premise and not on Exchange Server, enter the netbios name of an Exchange server. [This may not run through your load balancer for Kerberos authentication or if PS Remote is not enabled]  If it fails, you can copy this file to the Exchange Server and run as your Organizational Administrator.
File Default Location:  C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\PermissionsOnPremOrO365Combined.ps1

Log in as Global admin or Exchange Organizational Administrator.

Tip for Automatic Permissions:
Consider giving your Sync Service account this desired membership and you can set up Timed Automatic Permissions and give permissions with this account automatically

Global Admin or Exchange Organizational Administrator login Credentials:

When asked for service account, use your sync service account, default is ZADD2EXCHANGE

Select how to give permissions
Select option 0 to give to all, or option 3 to apply to only members of a distribution list

Distribution list names:
Do again until all are done

Again, consider using the automatic scheduled permissions option from the Diditbetter Support Menu

--------Off boarding--------

To off board users, remove the users from the attached Distribution list and then don’t hide, delete or detach or litigate or remove the Outlook license if 365 until at least 24 hours later.  This gives the program enough time to detect the user has been removed from the distribution list (Relman runs successfully) and then removes the data and the relationship automatically.

 
If done correctly, until at least 24 hours after, so the relationship can be marked and then auto dereplicate.  If the timing didn’t work, and the relationship for the *dead* account is in the trash and RED, you can select from Global Options to *Empty the Trash* and it close the Console and when prompted, you can select to *Not Continue where you left off* which will start at the beginning and remove all relationships in the trash, and force delete the relationships in the trash.
Let it run a complete cycle. If you open the Add2Exchange Console before the full sync, you will have to reselect and specify to *Empty the Trash* again.


Tips – If Outlook is installed for the profile, you must leave machine logged in (and locked if desired) as the sync service account user
Should be set to auto login.  This allows patch management to work correctly and continue syncing after a reboot.  For MS updates to be easily managed you can set this in the DidITBetter Support Menu.ps1 and the pw is encrypted. If the pw changes, run it again.  This auto login is primarily to recover after MS updates and a reboot.  Outlook does not need to be open for proper operation.

If using the automatic timed permissions option, use an account which password does not change, or give the service account proper permissions, since its password is set to never expire.

If in Office 365, consider making the service account a global admin.  If On Premise Exchange, consider making the service account an Exchange Organizational Administrator and set up Automatic Timed Permissions from the PowerShell *DidITBetter Support Menu.ps1* 


--------YTD--------


Some required and recommended finishing touches: 

If Contact Syncing (especially GAL) 

There is a default add in in Outlook add called the *Outlook Social Connector* which can cause the items in the destination to be changed.  
The effect is it is updating many of these contacts in every destination as evidenced by long relationship sync times.  A Ping pong effect occurs, where we sync to destination, the destination items change and then we change them back.  This is especially true if we make GAL entries as SMTP email address for Mobile utility. 
So for proper sync and put this continuing issue to rest, we immediately need to disable some or all parts of the Outlook Social Connector as group policy, and at least the option: *Blocking Global Address List synchronization with local contacts*. See MS link below 

After this policy is applied, it will require user logouts or reboots and back on of the machine or a forced group policy update, and closing and opening of Outlook.  After this, and two full syncs, we will be able to tell if other outside *off the domain* Outlook machines on an account, due to the longer times for the problem user.

Tip – if there are any Outlook clients on machines not logging into the domain (at home, for example) this will not get set by group policy, so should be disabled manually.  

It will also not be set if the Office 365 accounts are E1, and F1 accounts cannot be synced to as these are only online.  In order for Group Policy to be respected by the domain, they have to be Business or E3 accounts or better.  
Disabling can be done in the Add ins interface in Outlook, but often updates to MSO reenable those, so in the DidITBetter Support Menu.os1 (on the desktop) there is included the PowerShell script option to disable permanently by renaming the problem addins on machine – so that non domain type machine does not have the Social Connector enabled.  


How to Disable Outlook Social Connector by Group Policy
Our version: 
How to disable Outlook Social Connector by Group Policy or GP Edit
http://support.diditbetter.com/disable-group-policy.aspx

And Official Microsoft version: 
Manage the Outlook Social Connector with Group Policy
https://support.microsoft.com/en-us/help/2020103/how-to-manage-the-outlook-social-connector-by-using-group-policy


After this policy is applied it will require user logouts or reboots and back on of the machine, or a forced group policy update, and closing and opening of outlook.  After this, and two full syncs, we will be able to tell if other outside *off the domain* Outlook machines on an account, due to the longer times for that user.

How to Disable Outlook Social Connector by PowerShell on that machine

Copy the file from this location and run on the user’s machine
You may have to run as administrator and perhaps set this:
Set-executionpolicy -scope localmachine -executionpolicy unrestricted
This is the file, from the Replication server default location

C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Remove_Outlook_Add_ins.ps1

This PowerShell disables and renames the outlook addins so they are not accidentally reenabled.


Antivirus Exclusions from active file scanning
(These are default locations) 
C:\Program Files (x86)\OpenDoor Software®\
- Note the *Circle r* in the directory name. Copy and paste to antivirus program
C:\Program Files (x86)\DidItBetterSoftware
C:\Program Files\Microsoft SQL Server\
C:\Users\zadd2exchange\AppData
C:\Zlibrary
"

#History
LogWrite "--------History--------"
$History1 = Get-Content -Path "$Home\Desktop\Support.txt" -ErrorAction SilentlyContinue
LogWrite $History1 -ErrorAction SilentlyContinue

$History2 = Get-Content -Path "C:\zLibrary\Support.txt" -ErrorAction SilentlyContinue
LogWrite $History2 -ErrorAction SilentlyContinue

Start-Sleep -Seconds 2

Rename-Item -Path "$Home\Desktop\Support.txt" -NewName "Old_Support.txt" -ErrorAction SilentlyContinue
Rename-Item -Path "C:\zLibrary\Support.txt" -NewName "Old_Support.txt" -ErrorAction SilentlyContinue

Start-Sleep -Seconds 2

#Shortcut
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Support.lnk")
$Shortcut.TargetPath = "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Setup_Details.log"
$Shortcut.Save()

Write-Host "Done"
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting