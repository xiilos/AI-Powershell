<#
        .SYNOPSIS
        Microsoft Office Manual updater

        .DESCRIPTION
        Will start the process for Outlook to search for new updates

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

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #
Set-Location "C:\Program Files\Common Files\Microsoft Shared\ClickToRun"

.\OfficeC2RClient.exe /update user


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting