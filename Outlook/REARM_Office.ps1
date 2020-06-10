if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Set-Location "C:\Program Files (x86)\Microsoft Office\Office16"

.\OSPPREARM.EXE

cscript .\ospp.vbs /dstatus

Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting