if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Start-Process "https://downloads.malwarebytes.com/file/mb3" -Wait

Copy-Item "\\fileserv-db1\Sandbox\Port Apps\PortableApps\ccPortable\*" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools\ccPortable" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\Quick_Clean.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\Server 2016-2019 New Setup.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\VM_QuickUpdate_Clean.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\Windows 10 App Removal.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\Windows 10 New Setup.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse
Copy-Item "\\fileserv-db1\Work\Advantage International\PowerShell Tools REPO\AI-Powershell\Windows Services & Setup\Quick_Clean_v2.ps1" -Destination "\\fileserv-db1\Sandbox\Port Apps\Documents\Core Tools" -Recurse

Write-Host "Success"
Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting