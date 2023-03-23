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

# Specify the user's email address and profile name
$emailAddress = "zadd2exchange@yourdomain.com"
$profileName = "zadd2exchange"

# Create a new Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Create a new profile object
$profile = $outlook.Application.Profiles.Add($profileName)

# Create a new account object
$account = $profile.Accounts.Add()

# Set the account properties
$account.DisplayName = $profileName
$account.UserName = $emailAddress
$account.SmtpAddress = $emailAddress
$account.AccountType = [Microsoft.Office.Interop.Outlook.OlAccountType]::olExchange

# Save the changes to the profile and account
$profile.Save()
$account.Save()

# Open Outlook with the new profile
$outlook.Session.Logon($profileName)


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting