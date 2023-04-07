<#
        .SYNOPSIS
        

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

$wsusServerName = "YourWSUSServerName"
$wsusPortNumber = "8530" # Or use the default port number 80

$computerName = "YourComputerName"

# Generate a unique ID for the computer
$id = [guid]::NewGuid()

# Create a registry key to store the WSUS server configuration
$regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
New-Item -Path $regPath -Force | Out-Null

# Set the WSUS server and port number
Set-ItemProperty -Path $regPath -Name "WUServer" -Value "http://$wsusServerName:$wsusPortNumber"
Set-ItemProperty -Path $regPath -Name "WUStatusServer" -Value "http://$wsusServerName:$wsusPortNumber"
Set-ItemProperty -Path $regPath -Name "TargetGroup" -Value "Unassigned Computers"
Set-ItemProperty -Path $regPath -Name "TargetGroupEnabled" -Value 1

# Set the unique ID for the computer
Set-ItemProperty -Path $regPath -Name "ClientId" -Value $id.ToString()

# Trigger a detection cycle to register the computer with the WSUS server
$wu = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session"))
$wsus = $wu.CreateUpdateServiceManager()
$wsus.AddScanPackageService()
$wsus.ScanNow()


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting