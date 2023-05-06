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
$sourceEmailAddress = "sourceuser@domain.com"

# Set the distribution list's name
$distributionList = "Distribution List Name"

# Export contacts from the source user's mailbox
$contacts = Get-Contact -ResultSize Unlimited -Filter {EmailAddresses -like $sourceEmailAddress}

# Loop through the members of the distribution list and import the contacts to their mailboxes
Get-DistributionGroupMember -Identity $distributionList | ForEach-Object {
    $mailbox = $_.PrimarySmtpAddress
    foreach ($contact in $contacts) {
        New-MailContact -Name $contact.DisplayName -DisplayName $contact.DisplayName -ExternalEmailAddress $contact.EmailAddresses[0].Address -FirstName $contact.FirstName -LastName $contact.LastName -OrganizationalUnit "Contacts" -Alias $contact.Alias -PhoneNumber $contact.PhoneNumbers[0].PhoneNumber -MobilePhone $contact.MobilePhone -HomePhone $contact.HomePhone -Company $contact.Company -Department $contact.Department -StreetAddress $contact.StreetAddress -City $contact.City -StateOrProvince $contact.StateOrProvince -PostalCode $contact.PostalCode -CountryOrRegion $contact.CountryOrRegion -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $contact.EmailAddresses[0].Address -Confirm:$false -TargetAddress $mailbox
    }
}



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting