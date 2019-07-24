if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
Write-Host "Flushing DNS"
IPConfig /FlushDNS
Write-Host "Done"

Write-Host "Cleaning up Files via Disk Cleanup"
Cleanmgr /verylowdisk /SageRun:5 | Out-Null
Write-Host "Done"

Write-Host "Running CCleaner for Temp Files"
Push-Location ./ccPortable
Start-Process ccPortable.exe /AUTO -wait
Write-Host "Done"

Write-Host "Running CCleaner for the Registry"
Start-Process ccPortable.exe /Registry -wait
Write-Host "Done"

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting