if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
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
#Step 5: Add Registry Favs
#Step 6: Cleanup


# Start of Automated Scripting #

# Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1

# Office 365 and on premise Exchange Permissions


$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("In this step, we will assign the service account full mailbox access to the users that will be syncing with Add2Exchange. 
Instead of adding permissions to everyone, create a Distribution list with all users that will be syncing with Add2Exchange. 
Reminder*** You cannot hide this Distribution list, so it helps to put a Z in front of it to drop it to the bottom of the GAL",0,"Assigning Mailbox Permissions",0x1)
if($answer -eq 2){Break}

$message  = 'Please Pick how you want to connect'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange 2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&MExchange 2010'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Skip-Take me to Public Folder Permissions'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$choice = $Host.UI.PromptForChoice($message, $question, $choices, 4)

# Option 3: Skip

if ($choice -eq 3) {

Write-Host "Skipping"
}

# Option 4: Quit

if ($choice -eq 4) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


# Option 1: Office 365


if ($choice -eq 0) {

$error.clear()
Import-Module "MSonline" -ErrorAction SilentlyContinue
If($error){Write-Host "Adding Azure MSonline module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module MSonline -Confirm:$false -WarningAction "Inquire"} 
Else{Write-Host 'Module is installed'}


Write-Host "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

do {

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


$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Adding Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
Write-Host "Done"

}

# Option 4: Office 365-Remove Permissions within a dist. list

if ($decision -eq 4) {


$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Host "Removing Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done"

}

# Option 5: Office 365-Add Permissions to single user

if ($decision -eq 5) {


$Identity = read-host "Enter user Email Address"

Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
Write-Host "Done"

}

# Option 6: Office 365-Remove Single user permissions

if ($decision -eq 6) {


$Identity = read-host "Enter user Email Address"

Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"

}

# Option 7: Office 365-Quit

if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
  }

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange on Premise (2013-2016)


if ($choice -eq 1) {

$confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
if ($confirmation -eq 'y') {
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true
}

if ($confirmation -eq 'n') {
$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Before Continuing, please remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)
if($answer -eq 2){Break}
        
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true   
}


Write-Host "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
$User = read-host "Enter Sync Service Account";

Do {

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
Write-Host "Done"

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
Write-Host "Done"

}

# Option 3: Exchange on Premise-Adding Permissions to dist. list

if ($decision -eq 3) {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Adding Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
Write-Host "Done"
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
}
Write-Host "Done"

}

# Option 4: Exchange on Premise-Removing dist. list permissions

if ($decision -eq 4) {


$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Removing Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done"

}

# Option 5: Exchange on Premise-Adding permissions to single user

if ($decision -eq 5) {

$Identity = read-host "Enter user Email Address";

Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
Write-Host "Done"
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
}
Write-Host "Done"

}

# Option 6: Exchange on Premise-Removing permissions to single user

if ($decision -eq 6) {

$Identity = read-host "Enter user Email Address"
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"
}
  
# Option 7: Exchange on Premise-Quit
  
if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 2: Exchange 2010

if ($choice -eq 2) {

$confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
if ($confirmation -eq 'y') {
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010;
Set-ADServerSettings -ViewEntireForest $true
}
    
if ($confirmation -eq 'n') {
$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Before Continuing, please remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)
if($answer -eq 2){Break}
            
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true   
}
  
Write-Host "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
$User = read-host "Enter Sync Service Account";
  
do {

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
Write-Host "Done"
  
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
Write-Host "Done"
  
}
  
# Option 3: Exchange 2010 on Premise-Adding Permissions to dist. list
  
if ($decision -eq 3) {
  
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
  
Write-Host "Adding Add2Exchange Permissions"
  
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
Write-Host "Done"
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
}
Write-Host "Done"  
}
  
# Option 4: Exchange 2010 on Premise-Removing dist. list permissions
  
if ($decision -eq 4) {

$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
  
Write-Host "Removing Add2Exchange Permissions"
  
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done"
  
}
  
# Option 5: Exchange 2010 on Premise-Adding permissions to single user
  
if ($decision -eq 5) {
  
$Identity = read-host "Enter user Email Address";
  
Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
Write-Host "Done"
$confirmation = Read-Host "Would you like me add the A2E Throttling Policy? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
}
Write-Host "Done"   
}
  
# Option 6: Exchange 2010 on Premise-Removing permissions to single user
  
if ($decision -eq 6) {
  
$Identity = read-host "Enter user Email Address"
  
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"
}

# Option 7: Exchange 2010 on Premise-Quit
  
if ($decision -eq 7) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Public Folder Permissions

$confirmation = Read-Host "Do we need to add permissions to any Public Folders? [Y/N]"
if ($confirmation -eq 'N') {
Write-Host "Done"
Write-Host "Quitting"

}

if ($confirmation -eq 'Y') {
Write-Host "Taking you there now"
Get-Module | Remove-Module

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

Write-Host "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";

Do {

$message  = 'Please Pick what you want to do'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)



# Option 0: Office 365-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"

}

# Option 1: Office 365-Removing Public Folder Permissions

if ($decision -eq 1) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
  
}

# Option 2: Office 365-Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Option 1: Exchange on Premise-Public Folders (2013-2016)

if ($decision -eq 1) {
$confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
if ($confirmation -eq 'y') {
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Set-ADServerSettings -ViewEntireForest $true
}
    
if ($confirmation -eq 'n') {
$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Before Continuing, please remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)
if($answer -eq 2){Break}
            
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true   
 }

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";

Do {

$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 0: Exchange on Premise-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"

}

# Option 1: Exchange on Premise-Remove Public Folder Permissions

if ($decision -eq 1) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
 
    }
    
# Option 2: Exchange on Premise-Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
      }

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange 2010 on Premise Public Folders


if ($decision -eq 2) {

$confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
if ($confirmation -eq 'y') {
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
Set-ADServerSettings -ViewEntireForest $true
}
    
if ($confirmation -eq 'n') {
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Before Continuing, please $answer = remote into your Exchange server.
Open Powershell as administrator
Type: *Enable-PSRemoting* without the stars and hit enter.
Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)
if($answer -eq 2){Break}
            
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true   
}

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";

Do {
    
$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'
    
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))
    
$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)
    
    
# Option 0: Exchange 2010 on Premise-Adding Public Folder Permissions
    
if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
   
}
    
# Option 1: Exchange 2010 on Premise-Remove Public Folder Permissions
    
if ($decision -eq 1) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
       
}

# Option 2: Exchange 2010 on Premise- Quit
    
if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

}
  }

# Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3
# Auto Logon

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Excellent!! Permissions are done and now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue",0,"AutoLogin",0x1)
if($answer -eq 2){Break}

Start-Process -FilePath ".\Setup\AutoLogon.exe" -wait


# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4
# Installing the Software

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("System Setup Complete. Lets Install the Software",0,"Complete",0x1)
if($answer -eq 2){Break}

Start-Process -FilePath ".\Add2ExchangeSetup.msi" -wait -ErrorAction Stop



# Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5
# Registry Favorites & Shortcuts

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Once the Install is complete, Clcik OK to finish the setup",0,"Finishing Installation",0x1)
if($answer -eq 2){Break}

Write-Host "Creating Registry Favorites"
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Session Manager" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlNew\Control\Session Manager" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "EnableLUA" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name ".Net Framework" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "OpenDoor Software" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Add2Exchange"  -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Pendingfilerename" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\ControlNew001\Control\Session Manager" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "AutoDiscover" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Office" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Group Policy History" -Type string -Value "Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\History" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Logon" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Update" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -Force


# Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6
# Completion and Wrap-Up

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Setup is Complete. You can now start the Add2Echange Console",0,"Done",0x1)
if($answer -eq 2){Break}
Get-PSSession | Remove-PSSession
Exit
# End Scripting