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

# Import the Outlook COM object
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

# Create an instance of the Outlook application
$outlook = New-Object -ComObject Outlook.Application

# Get the Inbox folder
$inbox = $outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Get the last email in the Inbox
$email = $inbox.Items | Sort-Object -Property ReceivedTime | Select-Object -Last 1

# Check if the email subject contains a specific string
if ($email.Subject.Contains("Run PowerShell Script")) {

    # Run another PowerShell script
    & "C:\path\to\your\powershell\script.ps1"
}




<#This script does the following:

Imports the Outlook COM object to access Outlook functionality.
Creates an instance of the Outlook application.
Gets the Inbox folder.
Gets the last email in the Inbox.
Checks if the email subject contains a specific string.
If the email subject contains the specific string, it runs another PowerShell script located at the specified path.
Note that this script assumes that the user has Outlook installed and has already set up their email account in Outlook. 
Also, make sure to replace "Run PowerShell Script" and "C:\path\to\your\powershell\script.ps1" with the actual subject line string and PowerShell script path that you want to use.#>

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting