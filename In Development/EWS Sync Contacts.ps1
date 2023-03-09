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


# Load EWS Managed API assembly
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# Set source and destination mailbox credentials
$sourceCredential = Get-Credential -Message "Enter source mailbox credentials"
$destinationCredential = Get-Credential -Message "Enter destination mailbox credentials"

# Set source and destination mailbox email addresses
$sourceEmailAddress = "source@example.com"
$destinationEmailAddress = "destination@example.com"

# Set up EWS connection for source mailbox
$sourceExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$sourceExchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($sourceCredential.UserName, $sourceCredential.Password)
$sourceExchangeService.Url = New-Object System.Uri("https://outlook.office365.com/EWS/Exchange.asmx")

# Set up EWS connection for destination mailbox
$destinationExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$destinationExchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($destinationCredential.UserName, $destinationCredential.Password)
$destinationExchangeService.Url = New-Object System.Uri("https://outlook.office365.com/EWS/Exchange.asmx")

# Set up contacts folder for source mailbox
$contactsFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $sourceEmailAddress)
$contactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($sourceExchangeService, $contactsFolderId)

# Set up contacts folder for destination mailbox
$destinationContactsFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $destinationEmailAddress)
$destinationContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($destinationExchangeService, $destinationContactsFolderId)

# Get contacts from source mailbox
$contactsView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$contacts = $contactsFolder.FindItems($contactsView)

# Loop through contacts and create/update in destination mailbox
foreach ($contact in $contacts) {
    # Create new contact item in destination mailbox
    $newContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($destinationContactsFolder)
    $newContact.GivenName = $contact.GivenName
    $newContact.Surname = $contact.Surname
    $newContact.EmailAddresses[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1 = $contact.EmailAddresses[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1
    $newContact.PhoneNumbers[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone = $contact.PhoneNumbers[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone
    $newContact.PhoneNumbers[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone = $contact.PhoneNumbers[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone
    $newContact.Save()

    # Update existing contact item in destination mailbox
    $existingContact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($destinationExchangeService, $newContact.Id)
    $existingContact.GivenName = $contact.GivenName
    $existingContact.Surname = $contact.Surname
    $existingContact.EmailAddresses[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1 = $contact.EmailAddresses[



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting