if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Variables #

$User = "Type the Name of Sync Service account here"

# Script #

Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010;
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to All Users
Get-MailboxDatabase | Get-Mailbox -ResultSize Unlimited | Where-Object {$_.WhenCreated –ge ((Get-Date).Adddays(-1))} | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false



Get-PSSession | Remove-PSSession
Exit

# End Scripting