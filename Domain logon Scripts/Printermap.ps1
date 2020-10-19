if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Add-Printer -ConnectionName "\\DFSFILESERVER\HP LaserJet 4000 Series (Receptionist)" -ErrorAction SilentlyContinue
Add-Printer -ConnectionName "\\DFSFILESERVER\HP LaserJet 4000 Series (Accounting)" -ErrorAction SilentlyContinue
Add-Printer -ConnectionName "\\DFSFILESERVER\HP LaserJet P4014/P4015 (Office 1)" -ErrorAction SilentlyContinue

Get-PSSession | Remove-PSSession
Exit

# End Scripting