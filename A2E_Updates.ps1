if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted


Write-Host "Updating Add2Exchange"
Write-Host "Downloading the latest Add2Exchange"

Invoke-WebRequest "ftp://ftp.diditbetter.com/A2E-Enterprise/A2EAutoUpdate/Add2Exchange_Upgrade.msi" -outfile "c:\zlibrary\Add2Exchange_Upgrade.msi"

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting