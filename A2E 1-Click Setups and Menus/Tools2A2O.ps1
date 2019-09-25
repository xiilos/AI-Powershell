if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

Copy-Item "\\diditbetter\dev\DevLimAccess\a2e-enterprise\Contrib\Setup.zip" -Destination "\\diditbetter\dev\DevLimAccess\Add2Outlook\contrib\Support.zip" -Force

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting