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

# Replace the following variables with your own values:
$DistributionLists = @("DistributionList1@domain.com", "DistributionList2@domain.com")  # The list of distribution lists to grant access to
$UserToAdd = "zadd2exchange@domain.com"  # The user to grant full access to
$Domains = @("domain1.com", "domain2.com")  # The list of domains to grant access to


# Loop through each domain and distribution list and grant full access to the user for each user in the distribution list
foreach ($Domain in $Domains) {
    foreach ($DLName in $DistributionLists) {
        # Get the list of users in the distribution list
        $Users = Get-DistributionGroupMember -Identity $DLName -ResultSize unlimited | Where-Object { $_.RecipientType -eq "UserMailbox" }
        
        # Loop through each user and grant full access to the user
        foreach ($User in $Users) {
            $UserEmail = $User.PrimarySmtpAddress.ToString()
            Add-MailboxPermission -Identity $UserEmail -User $UserToAdd -AccessRights FullAccess -AutoMapping $false
            Write-Host "Granted full access to $UserEmail for $UserToAdd in $DLName and $Domain"
        }
    }
}





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting