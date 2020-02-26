if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Start-Process Powershell Read-Host "Local Account Password" -assecurestring | convertfrom-securestring | out-file "C:\password.txt"

pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting