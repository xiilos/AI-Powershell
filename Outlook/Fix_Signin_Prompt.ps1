if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Write-Host "Stopping Click to Run"
Stop-Service ClickToRunSvc -Force
Start-Sleep 5
Write-Host "Click to Run Service Stopped"
$Process = Get-Process -name "Outlook"
Write-Host "Restarting Outlook...."
Stop-Process $Process
Start-Sleep 3
Write-Host "Starting Outlook"
Start-Process Outlook


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting