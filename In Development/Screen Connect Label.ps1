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


reg query "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" /v "CurrentVersionDB"


({"label": "A2E Version", "command": reg query "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" /v "CurrentVersionDB"})


{"label": "Network Address", "command": "GuestNetworkAddress", "type": "builtin"})

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting