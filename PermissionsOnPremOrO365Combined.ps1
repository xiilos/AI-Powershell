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
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

if ($decision -eq 2) {
Exit
}





if ($decision -eq 0) {

Import-Module MSOnline

Write-Output "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic –AllowRedirection
Import-PSSession $Session
Import-Module MSOnline
Connect-MsolService –Credential $Cred


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


if ($decision -eq 0) {

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

Write-Output "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
Invoke-Item "C:\A2E_Office365_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 1) {

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

Write-Output "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess –verbose -confirm:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
Invoke-Item "C:\A2E_Office365_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 2) {

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

Write-Output "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess –verbose -confirm:$false
Write-Output "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
Invoke-Item "C:\A2E_Office365_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 3) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Output "Adding Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -AutoMapping:$false
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 4) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Output "Removing Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -Confirm:$false
}

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 5) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 6) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 7) {
Exit
}

}







if ($decision -eq 1) {


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


Set-ADServerSettings -ViewEntireForest $true


Write-Output "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
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


if ($decision -eq 0) {
  Write-Host 'Adding'
  Write-Output "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Output "Checking............"
Start-Sleep -s 2
Write-Output "All Done"
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
} 

if ($decision -eq 1) {
  Write-Host 'Removing'
  Write-Output "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess –verbose -Confirm:$false
Write-Output "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Output "Success....."
Write-Output "Checking............"
Start-Sleep -s 2
Write-Output "All Done"
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 2) {
    Write-Host 'Removing'
    Write-Output "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess –verbose -Confirm:$false
Write-Output "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Output "Success....."
Write-Output "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Output "Checking............"
Start-Sleep -s 2
Write-Output "All Done"
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 3) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Output "Adding Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Add-MailboxPermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -AutoMapping:$false
}
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 4) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$DistributionGroupName = read-host "Enter distribution list name (Display Name)";

Write-Output "Removing Add2Exchange Permissions"

$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -Confirm:$false
}
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 5) {

do {

$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Adding Add2Exchange Permissions to Single User"
Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 6) {
do {

$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Removing Add2Exchange Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 7) {
Exit
}
}