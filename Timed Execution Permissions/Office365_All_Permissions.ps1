if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$ServiceAccount = Get-Content ".\Setup\Timed Permissions\Creds\ServiceAccount.txt"
$Username = Get-Content ".\Setup\Timed Permissions\Creds\ServerUser.txt"
$Password = Get-Content ".\Setup\Timed Permissions\Creds\ServerPass.txt" | convertto-securestring

$Cred = New-Object -typename System.Management.Automation.PSCredential `
         -Argumentlist $Username, $Password

         Connect-MsolService -Credential $Cred
         $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
         Import-PSSession $Session -DisableNameChecking
         Import-Module MSOnline

#Timed Execution Permissions to All Users
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $ServiceAccount -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false

Get-PSSession | Remove-PSSession
Exit

# End Scripting