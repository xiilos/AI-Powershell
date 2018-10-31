If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}
#Execution Policy

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted

Get-AppxPackage | where-object {$_.name -notlike "*photos"} | where-object {$_.name -notlike "*store*"} | where-object {$_.name -notlike "*windowscalculator*"} | Remove-AppxPackage -Confirm:$False

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting