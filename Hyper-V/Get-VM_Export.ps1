if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Variables
#List All VM's Below
Get-VM –computername ‘Hyper-V01’,‘Hyper-V04’,‘Hyper-V05’,‘Hyper-V06’,‘Hyper-V08’,‘Hyper-V09’ | EXPORT-CSV C:\Current_VM.csv


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting