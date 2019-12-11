if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

$Server = "DIBEX16"

 

$URL = "ex16.diditbetter.com"

 

Get-OWAVirtualDirectory -Server $Server | Set-OWAVirtualDirectory -InternalURL "https://$($URL)/owa" -ExternalURL   "https://$($URL)/owa"

 

Get-ECPVirtualDirectory -Server $Server | Set-ECPVirtualDirectory -InternalURL "https://$($URL)/ecp" -ExternalURL   "https://$($URL)/ecp"

 

Get-OABVirtualDirectory -Server $Server | Set-OABVirtualDirectory -InternalURL "https://$($URL)/oab" -ExternalURL   "https://$($URL)/oab"

 

Get-ActiveSyncVirtualDirectory -Server $Server | Set-ActiveSyncVirtualDirectory -InternalURL "https://$($URL)/Microsoft-Server-ActiveSync" -ExternalURL "https://$($URL)/Microsoft-Server-ActiveSync"

 

Get-WebServicesVirtualDirectory -Server $Server | Set-WebServicesVirtualDirectory -InternalURL "https://$($URL)/EWS/Exchange.asmx" -ExternalURL "https://$($URL)/EWS/Exchange.asmx"

 

Get-MapiVirtualDirectory -Server $Server | Set-MapiVirtualDirectory -InternalURL "https://$($URL)/mapi" -ExternalURL https://$($URL)/mapi

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting