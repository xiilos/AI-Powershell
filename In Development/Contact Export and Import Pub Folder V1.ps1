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



# Specify the public folder and distribution list names
$PublicFolderName = "PublicFolderName"
$DistributionListName = "DistributionListName"

# Get all contacts from the public folder
$Contacts = Get-Contact -PublicFolder $PublicFolderName

# Get all users from the distribution list
$DistributionListMembers = Get-DistributionGroupMember -Identity $DistributionListName

# Loop through each user in the distribution list
foreach ($Member in $DistributionListMembers) {

    # Export contacts to a CSV file
    $Contacts | Export-Csv -Path "C:\Temp\Contacts.csv" -NoTypeInformation

    # Import contacts from the CSV file to the user's mailbox
    $UserEmailAddress = $Member.PrimarySmtpAddress
    Import-Csv -Path "C:\Temp\Contacts.csv" | ForEach-Object { New-MailContact -Name $_.Name -ExternalEmailAddress $_.EmailAddress -FirstName $_.FirstName -LastName $_.LastName -OrganizationalUnit $_.OrganizationalUnit -Alias $_.Alias -DisplayName $_.DisplayName -City $_.City -StateOrProvince $_.StateOrProvince -PostalCode $_.PostalCode -PhoneNumber $_.PhoneNumber -MobilePhone $_.MobilePhone -Fax $_.Fax -Company $_.Company -Department $_.Department -Title $_.Title -OtherTelephone $_.OtherTelephone -OtherMobile $_.OtherMobile -OtherFax $_.OtherFax -Notes $_.Notes -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $UserEmailAddress -ErrorAction SilentlyContinue }
}




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting