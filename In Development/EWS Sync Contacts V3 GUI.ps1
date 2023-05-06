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


# Create a GUI window
Add-Type -AssemblyName System.Windows.Forms
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Copy Contacts"
$Form.ClientSize = New-Object System.Drawing.Size(300,150)
$Form.StartPosition = "CenterScreen"

# Add labels and text boxes for input
$LabelSource = New-Object System.Windows.Forms.Label
$LabelSource.Text = "Source User Email:"
$LabelSource.AutoSize = $true
$LabelSource.Location = New-Object System.Drawing.Point(10,25)
$Form.Controls.Add($LabelSource)

$TextBoxSource = New-Object System.Windows.Forms.TextBox
$TextBoxSource.Location = New-Object System.Drawing.Point(120,25)
$TextBoxSource.Size = New-Object System.Drawing.Size(150,20)
$Form.Controls.Add($TextBoxSource)

$LabelTarget = New-Object System.Windows.Forms.Label
$LabelTarget.Text = "Target User Email:"
$LabelTarget.AutoSize = $true
$LabelTarget.Location = New-Object System.Drawing.Point(10,50)
$Form.Controls.Add($LabelTarget)

$TextBoxTarget = New-Object System.Windows.Forms.TextBox
$TextBoxTarget.Location = New-Object System.Drawing.Point(120,50)
$TextBoxTarget.Size = New-Object System.Drawing.Size(150,20)
$Form.Controls.Add($TextBoxTarget)

# Add a button to initiate the copy process
$ButtonCopy = New-Object System.Windows.Forms.Button
$ButtonCopy.Location = New-Object System.Drawing.Point(100,90)
$ButtonCopy.Size = New-Object System.Drawing.Size(100,30)
$ButtonCopy.Text = "Copy Contacts"
$Form.Controls.Add($ButtonCopy)

# Add an event handler for the button click
$ButtonCopy.Add_Click({
    # Get the source and target user email addresses
    $SourceEmail = $TextBoxSource.Text
    $TargetEmail = $TextBoxTarget.Text

    # Create the EWS service object with user credentials
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $Service.Credentials = Get-Credential

    # Set the Autodiscover URL and authenticate
    $Service.AutodiscoverUrl($SourceEmail)

    # Find the contacts folder in the source mailbox
    $ContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

    # Create a search filter to find all contacts
    $Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactSchema]::ItemClass,"IPM.Contact")

    # Define the view to retrieve all contacts in a single request
    $View = New-Object Microsoft.Exchange.WebServices.Data.ItemView([int]::MaxValue)

    # Find all contacts in the source folder
    $Contacts = $ContactsFolder.FindItems($Filter,$View)

    # Create the EWS service object for the target mailbox
    $TargetService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $TargetService.Credentials = Get-Credential

    # Set the Autodiscover URL and authenticate
    $TargetService.AutodiscoverUrl($TargetEmail)

# Get the contacts folder in the target mailbox
$TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($TargetService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

# Copy each contact to the target folder
foreach ($Contact in $Contacts) {
    $NewContact = $Contact.Clone()
    $NewContact.FolderId = $TargetFolder.Id
    $NewContact.Save()
}

# Show a message box indicating that the process is complete
[System.Windows.Forms.MessageBox]::Show("Contacts copied successfully.","Copy Contacts",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
})

$Form.ShowDialog() | Out-Null

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting






