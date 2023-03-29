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

# Replace <DistributionGroupName> with the name of the distribution group you want to mail enable
$DistributionGroupName = "<DistributionGroupName>"

# Replace <EmailAddress> with the email address you want to associate with the distribution group
$EmailAddress = "<EmailAddress>"

# Mail enable the distribution group
Enable-DistributionGroup -Identity $DistributionGroupName -Alias $DistributionGroupName -PrimarySmtpAddress $EmailAddress


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting