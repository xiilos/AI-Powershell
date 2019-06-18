if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Variables
$Domain = Read-Host "What is the name of your Domain? Example: DidItBetter.Com"

# Script #

#Test Path for Scripts
Write-Host "Creating Landing Zone"
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Scripts"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Scripts Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Scripts"
}

#Copy Files OSC_Disable
Write-Host "Copying Files Needed..."
Copy-Item ".\OSC_Disable.bat" -Destination "C:\Program Files (x86)\DidItBetterSoftware\Scripts" 

#Create the GPO
Import-GPO -BackupId 021E358C-8C08-414F-88E4-96F8A5A26C0A -TargetName "Outlook Social Connector" -path $PSScriptroot -CreateIfNeeded -Domain $Domain

Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting