if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Variables
$Time = Read-Host "Every How Many Days?"
$Location = Read-Host "Location of Working Directory? Type in Fulle Path"
$File = Read-Host "File Name"
$TaskName = Read-Host "Task Name?"
$Description = Read-Host "Description of Task?"


#Task
$Repeater = (New-TimeSpan -Days $Time)
$Duration = ([timeSpan]::maxvalue)
$Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument -NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file "$File"
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "$TaskName" -Description "$Description"
Write-Host "Done"

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting