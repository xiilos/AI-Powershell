<#
        .SYNOPSIS
        A2E Directory shortcut

        .DESCRIPTION
        Open Add2Exchange Directory


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


# Script #

$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
        Start-Process $Install

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting