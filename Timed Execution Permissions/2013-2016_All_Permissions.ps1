if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$Exchangename = Get-Content ".\Setup\Timed Permissions\Creds\Exchangename.txt"
$ServiceAccount = Get-Content ".\Setup\Timed Permissions\Creds\ServiceAccount.txt"
$Username = Get-Content ".\Setup\Timed Permissions\Creds\ServerUser.txt"
$Password = Get-Content ".\Setup\Timed Permissions\Creds\ServerPass.txt" | convertto-securestring

$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to All Users
Get-Mailbox -Resultsize Unlimited | Where-Object {$_.WhenCreated –ge ((Get-Date).Adddays(-1))} | Add-MailboxPermission -User $ServiceAccount -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false

Get-PSSession | Remove-PSSession
Exit

# End Scripting