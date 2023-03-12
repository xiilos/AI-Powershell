<#
        .SYNOPSIS
        A2E Export license and Profile 1 data

        .DESCRIPTION
        Exports A2E reg files license and profile 1 data
        places files in zlibrary

        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}
#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" C:\zlibrary\License_Info.Reg
REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" C:\zlibrary\Profile_1.Reg

Write-Host "Done"
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting