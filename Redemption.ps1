if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Unrestricted


Write-Output "Unregistering Old Redemption"
regsvr32 -u "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Console\Redemption.dll"
Start-Sleep -s 1
Write-Output "Removing old Redemption"
Remove-Item –path "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Console\Redemption.dll" –recurse
Start-Sleep -s 1
Write-Output "Registering new Redemption"
regsvr32 "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Service\Redemption.dll"


Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting