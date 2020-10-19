if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

New-PSDrive –Name "F" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\FileServer" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "L" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\Legal" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "P" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\HR" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "K" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\Marketing" –Persist  -ErrorAction SilentlyContinue
New-PSDrive –Name "X" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\Accounting" –Persist  -ErrorAction SilentlyContinue
New-PSDrive –Name "Y" –PSProvider FileSystem –Root "\\diditbetter.com\dfs" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "Q" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\Quickbooks" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "O" –PSProvider FileSystem –Root "\\diditbetter.com\dfs\Order Fulfillment" –Persist -ErrorAction SilentlyContinue


Start-Process Powershell "\\diditbetter\netlogon\PrinterMap.ps1"


Get-PSSession | Remove-PSSession
Exit

# End Scripting