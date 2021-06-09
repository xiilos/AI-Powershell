if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

schtasks /run /s prodc /tn "PS Windows Updates"
schtasks /run /s prodc2 /tn "PS Windows Updates"
schtasks /run /s ProDFS /tn "PS Windows Updates"
schtasks /run /s ProDHCP /tn "PS Windows Updates"
schtasks /run /s ProServer1 /tn "PS Windows Updates"
schtasks /run /s ProExchange16 /tn "PS Windows Updates"
schtasks /run /s ProExchange19 /tn "PS Windows Updates"
schtasks /run /s ProAdd2Exchange /tn "PS Windows Updates"

Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting