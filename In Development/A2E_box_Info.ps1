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

#Windows Version
$WindowsVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption

LogWrite "Windows Version= $WindowsVersion"

#PowerShell Version
$PShellVersion = $PSVersionTable.PSVersion

LogWrite "PowerShell Version= $PShellVersion"

#Database Version
$CurrentVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "CurrentVersionDB"

LogWrite "Current Version DB= $CurrentVersion"

#Database Instance
$DBInstance = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance"

LogWrite "DB Instance Name= $DBInstance"

#Database Server
$DBServer = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBServer"

LogWrite "DB Server Name= $DBServer"

#Install Location
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation"

LogWrite "Install Location= $Install"

#Server Name
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server"

LogWrite "Server Name= $ServerName"

#Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount"

LogWrite "Service Account Name= $ServiceAccount"

#License Address
$LicenseAddress = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseAddress"

LogWrite "License Address= $LicenseAddress"



Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting