<#
        .SYNOPSIS
        Disables Outlook Updates

        .DESCRIPTION
        Will disable automatic updates within outlook


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

Do {
  $confirmation = Read-Host "Would you like to Disable or Enable Outlook Updates [D/E]"
  if ($confirmation -eq 'D') {
    Write-Host "Disabling Outlook Updates"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled" -value False
    Write-Host "Done"
  }

    
  if ($confirmation -eq 'E') {
    Write-Host "Enabling Outlook Updates"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled" -value True
    Write-Host "Done"
  }
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting