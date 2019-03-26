if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:

  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Goal
# Assign Permissions for Add2Exchange

# Start of Automated Scripting #


$Title  = "How are We Connecting"
$Question = "Pick one of the following from below"

$0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Office 365"
$1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Exchange 2013-2016"
$2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Exchange 2010"
$3 = New-Object System.Management.Automation.Host.ChoiceDescription "&3-Skip-Take me to Public Folder Permissions"
$4 = New-Object System.Management.Automation.Host.ChoiceDescription "&4-Quit"

$Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2, $3, $4)
[int[]]$DefaultChoice = @(4)
$choice = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)


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

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Permissions to All Users"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Permissions to All Users"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Both-Remove/Add Permissions to All Users"
    $3 = New-Object System.Management.Automation.Host.ChoiceDescription "&3-Add Permissions to a Distribution List"
    $4 = New-Object System.Management.Automation.Host.ChoiceDescription "&4-Remove Permissions to a Distribution List"
    $5 = New-Object System.Management.Automation.Host.ChoiceDescription "&5-Add Permissions to a Single User"
    $6 = New-Object System.Management.Automation.Host.ChoiceDescription "&6-Remove ermissions to a Single User"
    $7 = New-Object System.Management.Automation.Host.ChoiceDescription "&7-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2, $3, $4, $5, $6, $7)
    [int[]]$DefaultChoice = @(7)
    $Decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)


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

    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Permissions to All Users"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Permissions to All Users"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Both-Remove/Add Permissions to All Users"
    $3 = New-Object System.Management.Automation.Host.ChoiceDescription "&3-Add Permissions to a Distribution List"
    $4 = New-Object System.Management.Automation.Host.ChoiceDescription "&4-Remove Permissions to a Distribution List"
    $5 = New-Object System.Management.Automation.Host.ChoiceDescription "&5-Add Permissions to a Single User"
    $6 = New-Object System.Management.Automation.Host.ChoiceDescription "&6-Remove ermissions to a Single User"
    $7 = New-Object System.Management.Automation.Host.ChoiceDescription "&7-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2, $3, $4, $5, $6, $7)
    [int[]]$DefaultChoice = @(7)
    $Decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)

#Exchange 2013-2016 Throttling Policy

$TestPath = Get-ThrottlingPolicy -identity A2EPolicy
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
Write-Host "A2E Policy Exists...Resuming"
   }

Else {   
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
   }


# Option 0: Exchange on Premise-Adding new permissions all

if ($decision -eq 0) {
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false



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

    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Permissions to All Users"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Permissions to All Users"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Both-Remove/Add Permissions to All Users"
    $3 = New-Object System.Management.Automation.Host.ChoiceDescription "&3-Add Permissions to a Distribution List"
    $4 = New-Object System.Management.Automation.Host.ChoiceDescription "&4-Remove Permissions to a Distribution List"
    $5 = New-Object System.Management.Automation.Host.ChoiceDescription "&5-Add Permissions to a Single User"
    $6 = New-Object System.Management.Automation.Host.ChoiceDescription "&6-Remove ermissions to a Single User"
    $7 = New-Object System.Management.Automation.Host.ChoiceDescription "&7-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2, $3, $4, $5, $6, $7)
    [int[]]$DefaultChoice = @(7)
    $Decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)
  
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

$Title  = "How are We Connecting"
$Question = "Pick one of the following from below"

$0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Office 365"
$1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Exchange 2013-2016"
$2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Exchange 2010"
$3 = New-Object System.Management.Automation.Host.ChoiceDescription "&3-Quit"

$Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2, $3)
[int[]]$DefaultChoice = @(3)
$Decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)

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

    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Public Folder Permissions"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Public Folder Permissions"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2)
    [int[]]$DefaultChoice = @(2)
    $decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)



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

    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Public Folder Permissions"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Public Folder Permissions"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2)
    [int[]]$DefaultChoice = @(2)
    $decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)

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
    
    $Title  = "Please Pick what you want to do"
    $Question = "Pick one of the following from below"
    
    $0 = New-Object System.Management.Automation.Host.ChoiceDescription "&0-Add Public Folder Permissions"
    $1 = New-Object System.Management.Automation.Host.ChoiceDescription "&1-Remove Public Folder Permissions"
    $2 = New-Object System.Management.Automation.Host.ChoiceDescription "&2-Quit"
    
    $Choices = [System.Management.Automation.Host.ChoiceDescription[]]($0, $1, $2)
    [int[]]$DefaultChoice = @(2)
    $decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, $DefaultChoice)
    
    
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

# End Scripting