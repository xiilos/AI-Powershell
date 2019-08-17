if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Variables #
$Username = Get-StoredCredential -target 'Exchange2016' -ascredentialobject  | Select-Object Username -ExpandProperty Username
$Password = Get-StoredCredential -target 'Exchange2016' | Select-Object password -ExpandProperty password
$ServiceAccount = Get-StoredCredential -target 'serviceaccount' | Select-Object Username -ExpandProperty Username
$ExchangeName = Get-StoredCredential -target 'Exchange2016' -ascredentialobject  | Select-Object targetname -ExpandProperty targetname

# Script #

$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $password


$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction SilentlyContinue
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $ServiceAccount -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false



pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting