if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

Do {
Write-Host "Changing and Resyncing Windows Time"
Write-Host "Stopping Windows Time"
net stop w32time
Write-Host "Changing Time Sync Pools to 0, 1, 2, & 3.us.pool.ntp.org"
w32tm /config /syncfromflags:manual /manualpeerlist:"0.us.pool.ntp.org,1.us.pool.ntp.org,2.us.pool.ntp.org,3.us.pool.ntp.org"
w32tm /config /reliable:yes
Write-Host "Success"
Write-Host "Starting Windows Time"
net start w32time
Write-Host "Checking..."
w32tm /query /configuration

$repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting