if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #
New-PSDrive –Name "L" –PSProvider FileSystem –Root "\\vsan.local\Local Virtual Machine Hard Drives" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "O" –PSProvider FileSystem –Root "\\fileserver-db1\OS_ISO" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "T" –PSProvider FileSystem –Root "\\Hyper-V10\D$" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "U" –PSProvider FileSystem –Root "\\Hyper-V11\D$" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "V" –PSProvider FileSystem –Root "\\Hyper-V12\D$" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "W" –PSProvider FileSystem –Root "\\Hyper-V13\D$" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "Y" –PSProvider FileSystem –Root "\\Hyper-V14\D$" –Persist -ErrorAction SilentlyContinue
New-PSDrive –Name "Z" –PSProvider FileSystem –Root "\\vsan.local\Veeam_Backups" –Persist -ErrorAction SilentlyContinue


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting