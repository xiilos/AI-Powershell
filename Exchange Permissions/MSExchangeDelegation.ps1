if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}




# Script #
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true

Get-ADUser -Properties msExchDelegateListBL,msExchDelegateListLink -LDAPFilter "(msExchDelegateListBL=*)" | Select-Object @{n='Distinguishedname';e={$_.distinguishedname}},@{n= 'alternate';e={$_.msExchDelegateListLink}} | Export-csv "c:\zlibrary\userlist.csv" –notypeinformation –noclobber

# End Scripting