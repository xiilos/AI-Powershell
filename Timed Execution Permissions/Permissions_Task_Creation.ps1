if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


# Script #

$Path = Read-Host "What is the path for your script? i.e. c:\zlibrary\Permissions.ps1"
$Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument -ExecutionPolicy Bypass -WindowStyle Hidden $Path

$Trigger =  New-ScheduledTaskTrigger -Daily -At 5am
Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting