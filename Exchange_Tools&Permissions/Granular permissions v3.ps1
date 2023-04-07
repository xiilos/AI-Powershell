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

# Define the distribution list name
$DistributionList = "DistributionListName"

# Define the folder name
$FolderName = "FolderName"

# Define the account name
$AccountName = "zadd2exchange"

# Get the distribution list members
$Members = Get-DistributionGroupMember $DistributionList

# Loop through each member and add the account as an owner of the folder
foreach ($Member in $Members) {
    Add-MailboxFolderPermission -Identity "$($Member.Alias):\Contacts\$FolderName" -User $AccountName -AccessRights Owner
}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting