if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";
do {
$Identity = read-host "Public Folder Name (Alias)"

#Public Folder Access Rights
Write-Host "Adding Permissions to Public Folders"
Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
Write-Host "Done"
$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting