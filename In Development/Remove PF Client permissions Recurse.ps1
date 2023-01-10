if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #


Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User Abel@goettsch.com -confirm:$false





$AllPublicFolders = Get-PublicFolder -Identity "\" -Recurse

ForEach ($PF in $Allpublicfolders) {
  Remove-PublicFolderClientPermission -Identity $PF.name -User Abel@goettsch.com -confirm:$false}



Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User Abel@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User darcange@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User earcangel@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User Enrique@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User Eric@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User jorge@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User linda@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User nate@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User regina@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User robg@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User sandy@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User sjensen@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User scott@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User talcala@goettsch.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User asander@wsa-usa.com -confirm:$false
Get-PublicFolder -Identity "\" -Recurse | Remove-PublicFolderClientPermission -User robg@wsa-usa.com -confirm:$false



Abel@goettsch.com
darcange@goettsch.com
earcangel@goettsch.com
Enrique@goettsch.com
Eric@goettsch.com
jorge@goettsch.com
linda@goettsch.com
nate@goettsch.com
regina@goettsch.com
robg@goettsch.com
sandy@goettsch.com
sjensen@goettsch.com
scott@goettsch.com
talcala@goettsch.com
asander@wsa-usa.com
robg@wsa-usa.com






Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting