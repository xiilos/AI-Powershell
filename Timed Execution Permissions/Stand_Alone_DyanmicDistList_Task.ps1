if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
# Report Correct File Path of DynamicDistribution List File
Write-Host "Creating Task"
  $Repeater = (New-TimeSpan -Minutes 720)
  $Duration = ([timeSpan]::maxvalue)
  $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
  $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file "ENTER FILE PATH HERE"'
  Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Exports a Dynamic Distribution List of Users into a Static List of Users"
  Write-Host "Done"


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting