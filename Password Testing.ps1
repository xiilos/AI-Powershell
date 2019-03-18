if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
$TestPath = "C:\zlibrary\Logs"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Log Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Logs"
}

$Logfile = "C:\zLibrary\Logs\A2E_Details.log"
Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}
# Script #

#Clear The Log
$confirmation = Read-Host "Would you like to clear the Log? [Y/N]"
if ($confirmation -eq 'N') {
    Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
  Clear-Content "C:\zLibrary\Logs\A2E_Details.log"
}


LogWrite "..............Add2Exchange Details.............."

# Script #
$Password = Get-ItemPropertyValue "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "serviceAccountPwd" | convertto-securestring

LogWrite "Service Account Passworde= $Password"

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting