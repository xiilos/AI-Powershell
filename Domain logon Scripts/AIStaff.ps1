if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

New-PSDrive –Name "F" –PSProvider FileSystem –Root "\\diditbetter.com\DFS\FileServer" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "Y" –PSProvider FileSystem –Root "\\diditbetter.com\DFS" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "O" –PSProvider FileSystem –Root "\\diditbetter.com\DFS\Order Fulfillment" –Persist -ErrorAction SilentlyContinue


Start-Process Powershell "\\diditbetter\netlogon\PrinterMap.ps1"


Get-PSSession | Remove-PSSession
Exit

# End Scripting