if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

$UserCredential = Get-Credential
If (!$UserCredential) { Exit }
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://A2E_O365/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting