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


# Load the EWS Managed API assembly
Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2
Install-Package Exchange.WebServices.Managed.Api
Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\Exchange.WebServices.Managed.Api.2.2.1.2\lib\net35\Microsoft.Exchange.WebServices.dll"




# Create a new Exchange service object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

# Set the credentials for the source mailbox
$sourceEmail = Read-Host "Enter the email address of the source mailbox"
$sourceCredential = Get-Credential -UserName $sourceEmail -Message "Enter the password for the source mailbox"
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($sourceCredential.UserName, $sourceCredential.Password)

# Set the URL for the EWS endpoint
$service.AutodiscoverUrl($sourceEmail)

# Get the contacts folder in the source mailbox
$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

# Set the credentials for the destination mailbox
$destEmail = Read-Host "Enter the email address of the destination mailbox"
$destCredential = Get-Credential -UserName $destEmail -Message "Enter the password for the destination mailbox"
$destService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$destService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($destCredential.UserName, $destCredential.Password)
$destService.AutodiscoverUrl($destEmail)

# Get the contacts folder in the destination mailbox
$destFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($destService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

# Get all the contacts in the source folder
$contacts = $folder.FindItems([Microsoft.Exchange.WebServices.Data.ItemView]::PageSizeMax)

# Create a progress bar for tracking the copy progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Minimum = 0
$progressBar.Maximum = $contacts.Count
$progressBar.Value = 0
$progressBar.Step = 1
$progressBar.Width = 400
$progressBar.Height = 30
$progressBar.Top = 50
$progressBar.Left = 50
$progressBar.Visible = $true

# Copy each contact to the destination folder
foreach ($contact in $contacts) {
    $newContact = $contact.Clone()
    $newContact.ParentFolderId = $destFolder.Id
    $newContact.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
    $progressBar.PerformStep()
}

# Hide the progress bar when the copy is complete
$progressBar.Visible = $false
Write-Host "All contacts copied to $destEmail mailbox."







Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting






