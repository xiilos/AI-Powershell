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




# Set the name of the public folder and distribution list
$PublicFolderName = "PublicFolderName"
$DistributionListName = "DistributionListName"

# Get the contacts from the public folder
$Contacts = Get-Contact -ResultSize Unlimited -PublicFolder $PublicFolderName

# Get the users from the distribution list
$DistributionList = Get-DistributionGroupMember -Identity $DistributionListName
$Users = $DistributionList | Where-Object {$_.RecipientType -eq "UserMailbox"}

# Loop through the users and add the contacts to their mailbox
ForEach ($User in $Users) {
    Write-Host "Adding contacts to $($User.PrimarySmtpAddress)..."
    $Contacts | ForEach-Object {
        $Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact($User.PrimarySmtpAddress)
        $Contact.GivenName = $_.FirstName
        $Contact.Surname = $_.LastName
        $Contact.EmailAddresses.Add($_.EmailAddress1)
        $Contact.Save()
    }
}




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting