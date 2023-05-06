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


# Import the EWS Managed API assembly
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# Set the credentials of the source and destination mailboxes
$sourceEmail = "sourceuser@domain.com"
$sourcePassword = "sourceuserpassword"
$destinationEmail = "destinationuser@domain.com"
$destinationPassword = "destinationuserpassword"

# Create a new ExchangeService object for the source mailbox
$sourceService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$sourceService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($sourceEmail, $sourcePassword)
$sourceService.AutodiscoverUrl($sourceEmail)

# Create a new ExchangeService object for the destination mailbox
$destinationService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$destinationService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($destinationEmail, $destinationPassword)
$destinationService.AutodiscoverUrl($destinationEmail)

# Define the source and destination folders to copy the contacts from and to
$sourceFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $sourceEmail)
$destinationFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $destinationEmail)

# Create a new ItemView object to retrieve all contacts from the source folder
$itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$contacts = $null
$moreContacts = $true
while ($moreContacts) {
    $contacts = $sourceService.FindItems($sourceFolderId, $itemView)
    if ($contacts.TotalCount -gt 0) {
        $destinationFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($destinationService, $destinationFolderId)
        foreach ($contact in $contacts.Items) {
            # Create a new Contact object and copy the properties from the source contact
            $newContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($destinationService)
            $newContact.Load($contact.PropertySet)
            $newContact.Save($destinationFolder.Id)
        }
        $itemView.Offset += $contacts.Items.Count
    }
    else {
        $moreContacts = $false
    }
}

Write-Host "Contacts copied successfully from $sourceEmail to $destinationEmail."





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting






