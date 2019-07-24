if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Variables #
$ExchangeUsername = Get-StoredCredential -target 'Exchange_Server' -ascredentialobject | Select-Object Username -ExpandProperty Username
$ExchangePassword = Get-StoredCredential -target 'Exchange_Server' | Select-Object password -ExpandProperty password
$SyncUsername = Get-StoredCredential -target 'Sync_Account' | Select-Object Username -ExpandProperty Username
$ExchangeServer = Get-StoredCredential -target 'Exchange_Server' -ascredentialobject | Select-Object Comment -ExpandProperty Comment

# Script #

$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Exchangeusername, $ExchangePassword

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to All Users
#Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $ServiceAccount -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
#Get-Mailbox -Resultsize Unlimited | Where-Object {$_.WhenCreated -ge ((Get-Date).Adddays(-1))} | Add-MailboxPermission -User $ServiceAccount -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false

Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $SyncUsername -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Get-Mailbox -Resultsize Unlimited | Where-Object { $_.WhenCreated -ge ((Get-Date).Adddays(-90)) } | Add-MailboxPermission -User $SyncUsername -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false



Get-PSSession | Remove-PSSession
Exit

# End Scripting