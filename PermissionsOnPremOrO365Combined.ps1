if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:

  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

$message  = 'Please Pick how you want to connect'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange 2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&MExchange 2010'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 3)

# Option 3: Quit

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
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

# Option 1: Office 365-Removing Add2Exchange Permissions

if ($decision -eq 1) {

Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -verbose -confirm:$false
Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

# Option 2: Office 365-Remove&Add Permissions

if ($decision -eq 2) {

Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -Verbose -confirm:$false
Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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
Get-PSSession | Remove-PSSession
Exit
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
Get-PSSession | Remove-PSSession
Exit
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
Get-PSSession | Remove-PSSession
Exit
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
Get-PSSession | Remove-PSSession
Exit
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


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
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
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
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

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

# Option 6: Exchange on Premise-Removing permissions to single user

if ($decision -eq 6) {
do {

$Identity = read-host "Enter user Email Address"

Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

# Option 7: Exchange on Premise- Quit

if ($decision -eq 7) {
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
}
}

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 2: Exchange 2010


if ($decision -eq 2) {


  add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010;
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
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
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
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
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
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
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
  
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
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
  
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
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
  
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
  }
  
  # Option 6: Exchange 2010 on Premise-Removing permissions to single user
  
  if ($decision -eq 6) {
  do {
  
  $Identity = read-host "Enter user Email Address"
  
  Write-Host "Removing Add2Exchange Permissions to Single User"
  Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
  
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'
  
  } Until ($repeat -eq 'n')
  
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
  }
  
  # Option 7: Exchange on Premise- Quit
  
  if ($decision -eq 7) {
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
  }
}


# End Scripting