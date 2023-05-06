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

#Check for MS Online Module
Write-Host "Checking for Exhange Online Module"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

IF (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "Exchange Online Module Exists"

    $InstalledEXOv2 = ((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()

    $LatestEXOv2 = (Find-Module -Name ExchangeOnlineManagement).Version.ToString()

    [PSCustomObject]@{
        Match = If ($InstalledEXOv2 -eq $LatestEXOv2) { Write-Host "You are on the latest Version" } 

        Else {
            Write-Host "Upgrading Modules..."
            Update-Module -Name ExchangeOnlineManagement -Force
            Write-Host "Success"
        }

    }


} 
Else {
    Write-Host "Module Does Not Exist"
    Write-Host "Downloading Exchange Online Management..."
    Install-Module –Name ExchangeOnlineManagement -Force
    Write-Host "Success"
}

Import-Module –Name ExchangeOnlineManagement -ErrorAction SilentlyContinue

Write-Host "Sign in to Office365 as Exchange Admin"

Connect-ExchangeOnline



# Set the source user's email address
$sourceUserEmail = "sourceuser@contoso.com"

# Set the distribution list's display name
$dlDisplayName = "DL-Name"

# Get the source user's contacts
$sourceContacts = Get-MailContact -ResultSize Unlimited -Filter {WindowsEmailAddress -eq $sourceUserEmail}

# Get the distribution list
$dl = Get-DistributionGroup -Filter {DisplayName -eq $dlDisplayName}

# Import the contacts to each member of the distribution list
foreach ($member in $dl.Members) {
    $targetUserEmail = Get-Mailbox $member | Select-Object WindowsEmailAddress
    foreach ($contact in $sourceContacts) {
        New-MailContact -Name $contact.DisplayName -DisplayName $contact.DisplayName -ExternalEmailAddress $contact.WindowsEmailAddress -FirstName $contact.FirstName -LastName $contact.LastName -OrganizationalUnit $targetUserEmail.OrganizationalUnit -Alias $contact.Alias -PhoneNumber $contact.Phone -MobilePhone $contact.MobilePhone -Fax $contact.Fax
    }
}




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting