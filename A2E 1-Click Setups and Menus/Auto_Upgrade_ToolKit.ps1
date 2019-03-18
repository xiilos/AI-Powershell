if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass

<#Remove Add2Outlook ToolKit

Write-Host "Removing A2O ToolKit"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Outlook ToolKit'"
$Program.Uninstall()
Write-Host "Done"
#>

#Create zLibrary\Add2Outlook ToolKit Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\A2OToolKit"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "A2O ToolKit Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\A2OToolKit"
}

#Downloading Add2Outlook ToolKit

Write-Host "Downloading Add2Outlook ToolKit"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/Add2Outlook%20ToolKit/Upgrades/Add2Outlook%20ToolKit%20Full%20Installation.exe"
$Output = "C:\zlibrary\A2OToolKit\Add2OutlookToolKitFullInstallation.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Outlook ToolKit

Write-Host "Unpacking Add2Outlook ToolKit"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\A2OToolKit"
Write-Host "Done"

#Installing Add2Outlook ToolKit

Write-Host "Installing Add2Outlook ToolKit"
Start-Process -FilePath "./Add2OutlookToolKitFullInstallation.exe" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 2

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting