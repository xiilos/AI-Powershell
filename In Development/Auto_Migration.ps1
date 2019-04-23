if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #


  $Service = Get-Service -Name Add2Exchange Service
  if ($Service.Status -ne "Running"){
  Start-Service $ServiceName
  Write-Host "Starting " $ServiceName " service" 
  " ---------------------- " 
  " Service is now started"
  }
  if ($arrService.Status -eq "running"){ 
  Write-Host "$ServiceName service is already started"
  }
  


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting