if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Generalizing VM Image
%WINDIR%\system32\sysprep\sysprep.exe /generalize /shutdown /oobe



Write-Host "Done"
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting