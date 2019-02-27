if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

function Show-Menu
{
     param (
           [string]$Title = 'Add2Exchange Permissions'
     )
     Clear-Host
     Write-Host "================ $Title ================"
     
     Write-Host "Please Pick how you want to connect"

     Write-Host "1: Press '1' for Office 365"
     Write-Host "2: Press '2' for Exchange 2010"
     Write-Host "3: Press '3' for Exchange 2013-2016"
     Write-Host "4: Press '4' for Public Folder Permissions"
     Write-Host "Q: Press 'Q' to Quit"

}




     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {

# Option 1: Office 365

'1' {
Clear-Host
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
}

do {
    Write-Host "Pick one of the following from below"

    Write-Host "1: Press '1' for Add Perm O365"
    Write-Host "2: Press '2' for Remove Perm O365"
    Write-Host "3: Press '3' for Both-Remove/Add"
    Write-Host "4: Press '4' for Add Dist-List Perm"
    Write-Host "4: Press '5' for Remove Dist-List Perm"
    Write-Host "4: Press '6' for Add Single Perm"
    Write-Host "4: Press '7' for Remove Single Perm"
    Write-Host "Q: Press 'Q' to Quit"

# Option 1: Office 365-Adding Add2Exchange Permissions
 '1' {
Clear-Host
Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Done"
}

# Option 2: Office 365-Removing Add2Exchange Permissions

 '2' {
Clear-Host
Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -verbose -confirm:$false
Write-Host "Done" 
}

# Option 3: Office 365-Remove&Add Permissions

'3' {
Clear-Host
Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -Verbose -confirm:$false
Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Done"
}

# Option 4: Office 365-Adding to Dist. List

'4' {
Clear-Host
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Adding Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}
Write-Host "Done"    
}

# Option 5: Office 365-Remove Permissions within a dist. list

'5' {
Clear-Host
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Removing Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done" 
}
# Option 6: Office 365-Add Permissions to single user

'6' {
Clear-Host
$Identity = read-host "Enter user Email Address"
Write-Host "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
Write-Host "Done"  
}

# Option 7: Office 365-Remove Single user permissions

'7' {
Clear-Host
$Identity = read-host "Enter user Email Address"
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"
}
    
# Option Q: Office 365-Quit
    
'Q'{ 
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

 } until ($input -eq 'q')


 
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 2: Exchange 2010

'2' {
Clear-Host
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
    Write-Host "Pick one of the following from below"

    Write-Host "1: Press '1' for Add Perm O365"
    Write-Host "2: Press '2' for Remove Perm O365"
    Write-Host "3: Press '3' for Both-Remove/Add"
    Write-Host "4: Press '4' for Add Dist-List Perm"
    Write-Host "4: Press '5' for Remove Dist-List Perm"
    Write-Host "4: Press '6' for Add Single Perm"
    Write-Host "4: Press '7' for Remove Single Perm"
    Write-Host "Q: Press 'Q' to Quit"

# Option 1: Exchange 2010 on Premise-Adding new permissions all
'1' {
Clear-Host
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "Done"      
}

# Option 2: Exchange 2010 on Premise-Remove old Add2Exchange permissions
  
'2' {
Clear-Host
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Done"
}

# Option 3: Exchange 2010 on Premise-Remove/Add Permissions all
  
'3' {
Clear-Host
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

# Option 4: Exchange 2010 on Premise-Adding Permissions to dist. list
  
'4' {
Clear-Host
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

# Option 5: Exchange 2010 on Premise-Removing dist. list permissions

'5' {
Clear-Host
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Removing Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done"
}

# Option 6: Exchange 2010 on Premise-Adding permissions to single user
  
'6' {
Clear-Host
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

# Option 7: Exchange 2010 on Premise-Removing permissions to single user
  
'7' {
Clear-Host
$Identity = read-host "Enter user Email Address"
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"
}

# Option Q: Exchange 2010 on Premise-Quit
    
'Q' {
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}
     
} until ($input -eq 'q')


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange on Premise (2013-2016)

'3' {
Clear-Host
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
    Write-Host "Pick one of the following from below"

    Write-Host "1: Press '1' for Add Perm O365"
    Write-Host "2: Press '2' for Remove Perm O365"
    Write-Host "3: Press '3' for Both-Remove/Add"
    Write-Host "4: Press '4' for Add Dist-List Perm"
    Write-Host "4: Press '5' for Remove Dist-List Perm"
    Write-Host "4: Press '6' for Add Single Perm"
    Write-Host "4: Press '7' for Remove Single Perm"
    Write-Host "Q: Press 'Q' to Quit"



# Option 1: Exchange on Premise-Adding new permissions all

'1' {
Clear-Host
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Adding Throttling Policy"
New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "Done" 
}

# Option 2: Exchange on Premise-Remove old Add2Exchange permissions
  
'2' {
Clear-Host
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Done"
}

# Option 3: Exchange on Premise-Remove/Add Permissions all

'3' {
Clear-Host
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

# Option 4: Exchange on Premise-Adding Permissions to dist. list

'4' {
Clear-Host
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

# Option 5: Exchange 2013-2016 on Premise-Removing dist. list permissions

'5' {
Clear-Host
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";
Write-Host "Removing Add2Exchange Permissions"
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
}
Write-Host "Done"
}

# Option 6: Exchange on Premise-Adding permissions to single user

'6' {
Clear-Host
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

# Option 7: Exchange 2013-2016 on Premise-Removing permissions to single user
  
'7' {
Clear-Host
$Identity = read-host "Enter user Email Address"
Write-Host "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Host "Done"
}
    
# Option Q: Exchange 2013-2016 on Premise-Quit
'Q' {
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}      
    } until ($input -eq 'q')

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

'4' {
Clear-Host
function Show-Menu
{
     param (
           [string]$Title = 'Add2Exchange Public Folder Permissions'
     )
     Clear-Host
     Write-Host "================ $Title ================"
     
     Write-Host "Please Pick how you want to connect"

     Write-Host "1: Press '1' for Office 365"
     Write-Host "2: Press '2' for Exchange 2010"
     Write-Host "3: Press '3' for Exchange 2013-2016"
     Write-Host "Q: Press 'Q' to Quit"
}





     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {

'1' {
Clear-Host
# Option 1: Office 365

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
        Write-Host "Pick one of the following from below"

        Write-Host "1: Press '1' for Add Perm O365"
        Write-Host "2: Press '2' for Remove Perm O365"
        Write-Host "Q: Press 'Q' to Quit"

# Option 1: Office 365-Adding Public Folder Permissions

'1' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
}

# Option 2: Office 365-Removing Public Folder Permissions

'2' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
}

Option Q: Office 365-Quit
    
'Q' {
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}
} until ($input -eq 'q')


 #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 1: Exchange 2010 on Premise Public Folders

# Option 2: Exchange 2010

'2' {
    Clear-Host
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
        Write-Host "Pick one of the following from below"
    
        Write-Host "1: Press '1' for Add Perm Exchange 2010"
        Write-Host "2: Press '2' for Remove Perm Exchange 2010"
        Write-Host "Q: Press 'Q' to Quit"

# Option 1: Exchange 2010 on Premise-Adding Public Folder Permissions

'1' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
}

# Option 2: Exchange 2010 on Premise-Remove Public Folder Permissions
    
'2' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
}

# Option Q: Exchange 2010 on Premise-Quit
    
'Q' {
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}
} until ($input -eq 'q')

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Option 3: Exchange on Premise-Public Folders (2013-2016)
'3' {
Clear-Host
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
    Write-Host "Pick one of the following from below"
    
    Write-Host "1: Press '1' for Add Perm Exchange 2013-2016"
    Write-Host "2: Press '2' for Remove Perm Exchange 2013-2016"
    Write-Host "Q: Press 'Q' to Quit"

# Option 1: Exchange on Premise-Adding Public Folder Permissions

'1' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
}


# Option 2: Exchange on Premise-Remove Public Folder Permissions

'2' {
Clear-Host
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
$Identity = read-host "Public Folder Name (Alias)"
Write-Host "Removing Permissions to Public Folders"
Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
Write-Host "Done"
}

# Option Q: Exchange 2013-2016 on Premise-Quit
'Q' {
Clear-Host
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}      
} until ($input -eq 'q')



# End Scripting