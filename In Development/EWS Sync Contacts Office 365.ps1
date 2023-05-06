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


# Import the EWS Managed API module
Import-Module "C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll"

# Set up credentials and connection to source mailbox
$sourceEmailAddress = Read-Host "Enter the email address of the source mailbox"
$sourcePassword = Get-Credential -UserName $sourceEmailAddressail -Message "Enter the password for the source mailbox"
$sourceCredential = New-Object System.Management.Automation.PSCredential ($sourceEmailAddress, $sourcePassword)
$sourceService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$sourceService.Url = New-Object Uri("https://outlook.office365.com/EWS/Exchange.asmx")
$sourceService.Credentials = $sourceCredential

# Set up credentials and connection to destination mailbox
$destinationEmailAddress = "user 2"
$destinationPassword = ConvertTo-SecureString "password" -AsPlainText -Force
$destinationCredential = New-Object System.Management.Automation.PSCredential ($destinationEmailAddress, $destinationPassword)
$destinationService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$destinationService.Url = New-Object Uri("https://outlook.office365.com/EWS/Exchange.asmx")
$destinationService.Credentials = $destinationCredential

# Define the source and destination mailbox folders
$sourceFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$sourceEmailAddress)
$destinationFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$destinationEmailAddress)

# Set up a search filter for all items in the source folder
$searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactSchema]::DisplayName, "A*")
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$searchResults = $null
$sourceContacts = @()

# Retrieve all contacts in the source folder
do {
    $searchResults = $sourceService.FindItems($sourceFolderId, $searchFilter, $view)
    foreach ($item in $searchResults.Items) {
        $sourceContacts += $item
    }
    $view.Offset += $searchResults.Items.Count
} while ($searchResults.MoreAvailable -eq $true)

# Copy each contact to the destination folder
foreach ($contact in $sourceContacts) {
    $newContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($destinationService)
    $newContact.Copy($contact.Id, $destinationFolderId)
}

Write-Host "Contacts copied successfully!"




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting






