if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Goal
# Create Initial Environment for Add2Exchange Install
# 7 Steps

#Step 1: Account Creation
#Step 2: Upgrade .Net and Powershell if needed
#Step 3: Create zLibrary and Download Add2Exchange Software
#Step 4: Install Outlook and Setup Profile
#Step 5: Mailbox Creation
#Step 6: Create a Mail Profile
#Step 7: Disable UAC


# Start of Automated Scripting #

#Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1
# Account Creation

$message  = 'Have you created an account for Add2Exchange?'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 2: Quit

if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


if ($decision -eq 0) {


# Account Creation
Write-host "We need to create an account for Add2Exchange"
$confirmation = Read-Host "Are you on a domain? [Y/N]"
if ($confirmation -eq 'y') {

$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please create a new account called zadd2exchange (or any name of your choosing).
Note* The account only needs Domain User and Public Folder management permissions.
Note* You cannot hide this account.
Once Done, add the new user account as a local Administrator of this box.
Log off and back on as the new Sync Service account and run this again before you Proceed.",0,"Account Creation",0x1)
$confirmation = Read-host "Do you want to continue [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Resumming"
}

if ($confirmation -eq 'n') {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }
}



if ($confirmation -eq 'n') {
Write-Host "Lets create a new user for Add2Exchange"
$Username = Read-Host "Username for new account"
$Fullname = Read-host "User Full Name"    
$password = Read-Host "Password for new account" -AsSecureString
$description = read-host "Description of this account."

    New-LocalUser "$Username" -Password $Password -FullName "$Fullname" -Description "$description"
    Add-LocalGroupMember -Group "Administrators" -Member $Username

}
}

if ($decision -eq 1) {
    $confirmation = Read-Host "Are you logged into this box as the new sync account? [Y/N]"
    if ($confirmation -eq 'y') {
    Write-Host "Resumming"
    }

    
    if ($confirmation -eq 'n') {  
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Add the new user account as a local Administrator of this box.
    Log off and back on as the new Sync Service account and run this again before proceeding.",0,"Account Creation",0x1)
    }
}

# Step 2-----------------------------------------------------------------------------------------------------------------------------------------------------Step 2

# Powershell Update
# Check if .Net 4.5 or above is installed

$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389))
{
Write-Host "We need to download .Net 4.5.2"
Write-Host "Downloading"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}

$url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
$output = "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
(New-Object System.Net.WebClient).DownloadFile($url, $output)    
Start-Process -FilePath "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


#Check Operating Sysetm

$BuildVersion = [System.Environment]::OSVersion.Version


#OS is 10+
if($BuildVersion.Major -like '10')
    {
        Write-Host "WMF 5.1 is not supported for Windows 10 and above"
        
    }

#OS is 7
if($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1')
    {
        
Write-Host "Downloading WMF 5.1 for 7+"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}
$url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
$output = "C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Process -FilePath 'C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu' -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

#OS is 8
elseif($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3')
    {
        
Write-Host "Downloading WMF 5.1 for 8+"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}
$url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
$output = "C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Process -FilePath 'C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu' -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("PPlease Reboot after Installing and run this again",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

Write-Host "Nothing to do"
Write-Host "You Are on the latest version of PowerShell"


#Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3

#Create zLibrary

$TestPath = "C:\zlibrary"

if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "zLibrary exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path C:\zlibrary
 }

#Download the latest Full Installation
$confirmation = Read-Host "Would you like me to download the latest Add2Exchange? [Y/N]"
    if ($confirmation -eq 'y') {

        Write-Host "Downloading Add2Exchange......"
        Write-Host "This can take a few Minutes"
        $url = "ftp://ftp.diditbetter.com/A2E-Enterprise/A2ENewInstall/a2e-enterprise.zip"
        $output = "C:\zlibrary\a2e-enterprise.zip"
        (New-Object System.Net.WebClient).DownloadFile($url, $output)
        Write-Host "Download Done"
        Write-Host "Unpacking"    
        Expand-Archive -Path "C:\zlibrary\a2e-enterprise.zip" -DestinationPath "c:\zlibrary"

}



# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4

# Outlook Install 32Bit

If (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application -ErrorAction SilentlyContinue) {
    Write-Output "Outlook Is Installed"
}
Else {
Write-host "We need to install Outlook on this machine"
$confirmation = Read-Host "Would you like me to Install Outlook 365? [Y/N]"
    if ($confirmation -eq 'y') {

Start-Process -filepath "C:\zLibrary\A2E-Enterprise\O365Outlook32\OfficeProPlus_x32.msi" -wait

}

if ($confirmation -eq 'n') {

$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("If you choose to install your own Outlook; ensure that it is Outlook 2016 and above, and 32Bit Only.
When Done, click OK",0,"Outlook Install",0x1)
    
    }
}


# Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5

# Mailbox Creation

$confirmation = Read-Host "Are you on Office 365 or Exchange on Premise [O/E]"
if ($confirmation -eq 'O') {
Write-host "Make sure to create a mailbox for the sync service account and add an E3 License to it"
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("When this is Done, click OK",0,"Create a Mailbox",0x1)
}

if ($confirmation -eq 'E') {
Write-Host "Create a Mailboix for the new Sync Service account in Exchange"
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("When this is Done, Click OK",0,"Create a Mailbox",0x1)
}

# Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6

# Mail Profile

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("The next step is to Create a Profile for your new account. Open Control panel and go to Mail. Create a new profile and follow through the steps that pertain to your Organization. 
Note* Make sure you do not have Cache checked. When this is finished click OK",0,"Creating an Outlook Profile",0x1)

# Step 7-----------------------------------------------------------------------------------------------------------------------------------------------------Step 7

# Disable UAC
Write-Host "Checking UAC"

$Val = Get-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA"

if($val.EnableLUA -ne 0)

{
Set-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA" -value 0
Write-Host "Done"
Write-Host "UAC is now Disabled"
$wshell = New-Object -ComObject Wscript.Shell
    
$wshell.Popup("Initinal Setup is Complete. Please Reboot and run 1 click install",0,"Done",0x1)
  
}

Else {

Write-Host "UAC is already Disabled"
Write-Host "Resuming"
Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
    
$wshell.Popup("Initinal Setup is Complete. Please run 1 click install",0,"Done",0x1)
    
}

#Goal
# Assign Permissions for Add2Exchange
# Install Add2Exchange
# Cleanup
# Luanch Add2Exchange for the first time

#Step 1: Add Permissions
#Strp 2: Add Public Folder Permissions
#Step 3: Enable AutoLogon
#Step 4: Install Add2Exchange
#Step 5: Add Registry Favs and Shortcuts
#Step 6: Cleanup


# Start of Automated Scripting #

# Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1

# Office 365 and on premise Exchange Permissions


$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("In this step, we will assign the service account full mailbox access to the users that will be syncing with Add2Exchange. 
Instead of adding permissions to everyone, create a Distribution list with all users that will be syncing with Add2Exchange. 
Reminder*** You cannot hide this Distribution list, so it helps to put a Z in front of it to drop it to the bottom of the GAL",0,"Assigning Mailbox Permissions",0x1)

$message  = 'Please Pick how you want to connect'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange 2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&MExchange 2010'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Skip'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 4)

# Option 3: Skip

if ($decision -eq 3) {

Write-Host "Skipping"
}

# Option 4: Quit

if ($decision -eq 4) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


# Option 1: Office 365


if ($decision -eq 0) {

$error.clear()
Import-Module "MSonline" -ErrorAction SilentlyContinue
If($error){Write-Host "Adding Azure MSonline module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module MSonline -Confirm:$false -WarningAction "Inquire"} 
Else{Write-Host 'Module is installed'}

Import-Module MSOnline

Write-Host "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session
Import-Module MSOnline
Connect-MsolService -Credential $Cred -ErrorAction "Inquire"


$message  = 'Please Pick what you want to do'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Add Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Remove Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&5 Add Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&6 Remove Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&7 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 7)

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

# Option 0: Office 365-Adding Add2Exchange Permissions

if ($decision -eq 0) {

Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Done"

}

# Option 1: Office 365-Removing Add2Exchange Permissions

if ($decision -eq 1) {

Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -verbose -confirm:$false
Write-Host "Done"

}

# Option 2: Office 365-Remove&Add Permissions

if ($decision -eq 2) {

Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -Verbose -confirm:$false
Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Done"

}

# Option 3: Office 365-Adding to Dist. List

if ($decision -eq 3) {

do {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Adding Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"

}

# Option 4: Office 365-Remove Permissions within a dist. list

if ($decision -eq 4) {

do {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Removing Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"

}

# Option 5: Office 365-Add Permissions to single user

if ($decision -eq 5) {

do {

$Identity = read-host "Enter user Email Address"

Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"

}

# Option 6: Office 365-Remove Single user permissions

if ($decision -eq 6) {

do {

$Identity = read-host "Enter user Email Address"

Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"

}

# Option 7: Office 365-Quit

if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
  }

}

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange on Premise


if ($decision -eq 1) {

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Before Continuing, please remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)

$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

Write-Host "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
$User = read-host "Enter Sync Service Account";

$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Add Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Remove Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&5 Add Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&6 Remove Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&7 Quit'))



$decision = $Host.UI.PromptForChoice($message, $question, $choices, 7)

# Option 0: Exchange on Premise-Adding new permissions all

if ($decision -eq 0) {
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "All Done"

} 

# Option 1: Exchange on Premise-Remove old Add2Exchange permissions

if ($decision -eq 1) {
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Done"

}

# Option 2: Exchange on Premise-Remove/Add Permissions all

if ($decision -eq 2) {
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -Verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Success....."
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Checking............"
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
Write-Host "All Done"

}

# Option 3: Exchange on Premise-Adding Permissions to dist. list

if ($decision -eq 3) {

do {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Adding Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')


}

# Option 4: Exchange on Premise-Removing dist. list permissions

if ($decision -eq 4) {

do {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Removing Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')


}

# Option 5: Exchange on Premise-Adding permissions to single user

if ($decision -eq 5) {

do {

$Identity = read-host "Enter user Email Address";

Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')


}

# Option 6: Exchange on Premise-Removing permissions to single user

if ($decision -eq 6) {
do {

$Identity = read-host "Enter user Email Address"

Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')


}
  
# Option 7: Exchange on Premise-Quit
  
if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }
}

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 2: Exchange 2010


if ($decision -eq 2) {

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Before Continuing, please remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)

$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true
  
Write-Host "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
$User = read-host "Enter Sync Service Account";
  
$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'
  
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Add Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Remove Dist-List Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&5 Add Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&6 Remove Single Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&7 Quit'))
  
  
$decision = $Host.UI.PromptForChoice($message, $question, $choices, 7)
  
# Option 0: Exchange 2010 on Premise-Adding new permissions all
  
if ($decision -eq 0) {
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "All Done"
  
} 
  
# Option 1: Exchange 2010 on Premise-Remove old Add2Exchange permissions
  
if ($decision -eq 1) {
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Done"
  
}
  
# Option 2: Exchange 2010 on Premise-Remove/Add Permissions all
  
if ($decision -eq 2) {
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -Verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Success....."
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Checking............"
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
Write-Host "All Done"
  
}
  
# Option 3: Exchange 2010 on Premise-Adding Permissions to dist. list
  
if ($decision -eq 3) {
  do {
  
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
  
Write-Host "Adding Add2Exchange Permissions"
  
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
}
  
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
  
} Until ($repeat -eq 'n')
  
}
  
# Option 4: Exchange 2010 on Premise-Removing dist. list permissions
  
if ($decision -eq 4) {
 do {  

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
  
Write-Host "Removing Add2Exchange Permissions"
  
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
  
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
  
} Until ($repeat -eq 'n')  
  
}
  
# Option 5: Exchange 2010 on Premise-Adding permissions to single user
  
if ($decision -eq 5) {
  
do {
  
$Identity = read-host "Enter user Email Address";
  
Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
}
  
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
  
} Until ($repeat -eq 'n')
   
}
  
# Option 6: Exchange 2010 on Premise-Removing permissions to single user
  
if ($decision -eq 6) {
do {
  
$Identity = read-host "Enter user Email Address"
  
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
  
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
 
} Until ($repeat -eq 'n')
    
}
  
# Option 7: Exchange 2010 on Premise-Quit
  
if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }
}

# Step 2-----------------------------------------------------------------------------------------------------------------------------------------------------Step 2
#Public Folder Permissions

$confirmation = Read-Host "Do we need to add permissions to any Public Folders? [Y/N]"
if ($confirmation -eq 'N') {
Write-host "Resuming"

}

if ($confirmation -eq 'Y') {
Write-Host "Taking you there now"


$message  = 'Please Pick how you want to connect'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&MExchange2010'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 3)

# Option 2: Quit

if ($decision -eq 3) {

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


# Option 1: Office 365


if ($decision -eq 0) {

$error.clear()
Import-Module "MSonline" -ErrorAction SilentlyContinue
If($error){Write-Host "Adding Azure MSonline module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module MSonline -Confirm:$false -WarningAction "Inquire"} 
Else{Write-Host 'Module is installed'}

Import-Module MSOnline

Write-Host "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session
Import-Module MSOnline
Connect-MsolService -Credential $Cred -ErrorAction "Inquire"


$message  = 'Please Pick what you want to do'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";


# Option 0: Office 365-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
    do {
    $Identity = read-host "Public Folder Name (Alias)"
    Write-Host "Adding Permissions to Public Folders"
    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
    Write-Host "Done"
    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
    
    } Until ($repeat -eq 'n')

}

# Option 1: Office 365-Removing Public Folder Permissions

if ($decision -eq 1) {
    Write-Host "Getting a list of Public Folders"
    Get-PublicFolder -Identity "\" -Recurse
        do {
        $Identity = read-host "Public Folder Name (Alias)"
        Write-Host "Removing Permissions to Public Folders"
        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
        Write-Host "Done"
        $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        
        } Until ($repeat -eq 'n')
  
}
    
# Option 2: Office 365-Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
      }

}
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Option 1: Exchange on Premise


if ($decision -eq 1) {

$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 0: Exchange on Premise-Adding Public Folder Permissions

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";


# Option 0: Exchange on Premise-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
do {
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
    
} Until ($repeat -eq 'n')

}

# Option 1: Exchange on Premise-Remove Public Folder Permissions

if ($decision -eq 1) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
do {
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
        
} Until ($repeat -eq 'n')
 
    }


    
# Option 2: Exchange on Premise-Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
      }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange 2010 on Premise


if ($decision -eq 2) {

    
$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'
    
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))
    
$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)
    
# Option 0: Exchange 2010 on Premise-Adding Public Folder Permissions
    
#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";
    
    
# Option 0: Exchange 2010 on Premise-Adding Public Folder Permissions
    
if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
do {
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
        
} Until ($repeat -eq 'n')
   
}
    
# Option 1: Exchange 2010 on Premise-Remove Public Folder Permissions
    
if ($decision -eq 1) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
do {
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'
            
} Until ($repeat -eq 'n')
       
}
    
# Option 2: Exchange 2010 on Premise- Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}
}
  }

# Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3
# Auto Logon

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Excellent!! Permissions are done and now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue",0,"AutoLogin",0x1)

Start-Process -FilePath "C:\zlibrary\A2E-Enterprise\Setup\AutoLogon.exe" -wait -ErrorAction Stop


# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4
# Installing the Software

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("System Setup Complete. Lets Install the Software",0,"Complete",0x1)

Start-Process -FilePath "C:\zlibrary\A2E-Enterprise\Add2ExchangeSetup.msi" -wait -ErrorAction Stop



# Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5
# Registry Favorites & Shortcuts

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Once the Install is complete, Clcik OK to finish the setup",0,"Finishing Installation",0x1)

Write-Host "Creating Registry Favorites"
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Session Manager" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "EnableLUA" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name ".Net Framework" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "OpenDoor Software" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Add2Exchange"  -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Pendingfilerename" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Session Manager" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "AutoDiscover" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Office" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Group Policy History" -Type string -Value "Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\History" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Logon" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Update" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" | out-null

# Desktop Shortcuts

$wshShell = New-Object -ComObject "WScript.Shell"
$urlShortcut = $wshShell.CreateShortcut(
(Join-Path $wshShell.SpecialFolders.Item("AllUsersDesktop") "Disable Outlook Social Connector through GPO.url")
)
$urlShortcut.TargetPath = "https://support.microsoft.com/en-us/help/2020103/how-to-manage-the-outlook-social-connector-by-using-group-policy"
$urlShortcut.Save()

Copy-Item -Path "C:\zlibrary\A2E-Enterprise\support.txt" -Destination "$home\Desktop\Support.txt"

# Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6
# Completion and Wrap-Up

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Setup is Complete. You can now start the Add2Echange Console",0,"Done",0x1)
Get-PSSession | Remove-PSSession
Exit
# End Scripting