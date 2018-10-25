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
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Granular Perm to Single User'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Granular Perm to Single User'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)




if ($decision -eq 0) {


do {
$User = read-host "Enter Sync Service Account";
$Mailbox = read-host "(Enter user Email Address)"
$Identity = read-host "Enter user Email Address followed by :\Calendar or :\Contact (Example: Tom@diditbetter.com:\Contacts)"
$AccessRights = read-host "Enter Permissions level (Owner, Editor, none)"

Write-Output "Adding Add2Outlook Granular Permissions to Single User"

Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights 'readpermission'
Add-MailboxFolderPermission -Identity $Identity -User $User -AccessRights $AccessRights

  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')


Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


if ($decision -eq 1) {

do {
$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Removing Add2Outlook Granular Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 2) {
Exit
}

}







if ($decision -eq 1) {


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


Set-ADServerSettings -ViewEntireForest $true




$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Granular Perm to Single User'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Granular Perm to Single User'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))



$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)




if ($decision -eq 0) {
do {
Write-Output "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Outlook or zAdd2Outlook@yourdomain.com"

$User = read-host "Enter Sync Service Account";
$Mailbox = read-host "(Enter user Email Address)"
$Identity = read-host "Enter user Email Address followed by :\Calendar or :\Contact (Example: Tom@diditbetter.com:\Contacts)"
$AccessRights = read-host "Enter Permissions level (Owner, Editor, none)"

Write-Output "Adding Add2Outlook Granular Permissions to Single User"

Add-MailboxPermission -Identity $Mailbox -User $User  -AccessRights 'readpermission'
Add-MailboxFolderPermission -Identity $Identity -User $User -AccessRights $AccessRights

Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2O_permissions.txt
Invoke-Item "C:\A2O_permissions.txt"
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 1) {
do {
$User = read-host "Enter Sync Service Account (Display Name)";
$Identity = read-host "Enter user Email Address"

Write-Output "Removing Add2Outlook Granular Permissions to Single User"
Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
Write-Output "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2O_permissions.txt
Invoke-Item "C:\A2O_permissions.txt"
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

if ($decision -eq 2) {
Exit
}
}