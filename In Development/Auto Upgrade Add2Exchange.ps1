if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass


#Stop Add2Exchange Service

Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-sleep -s 5
Write-Host "Done"

#Remove Add2Exchange

Write-Host "Removing Add2Exchange"
Write-Hist "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "zLibrary exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path C:\zlibrary
 }

#Downloading Add2Exchange

Write-Host "Downloading Add2Exchange; Please wait...."
Invoke-WebRequest "ftp://ftp.diditbetter.com/A2E-Enterprise/Upgrades/a2e-enterprise_upgrade.exe" -outfile "c:\zlibrary\a2e-enterprise_upgrade.exe"
Write-Host "Done"

#Unpacking Add2Exchange

Write-Host "Unpacking Add2exchange"
Write-Host "please Wait....."
Set-Location c:\zlibrary
Start-Process "c:\zlibrary\a2e-enterprise_upgrade.exe" -wait
Write-Host "Done"

#Installing Add2Exchange

Write-Host "Installing Add2Exchange"
$Location = Get-ChildItem -Path . -Recurse | Where-Object {$_.LastWriteTime -gt (Get-Date).AddSeconds(-30)}
Set-Location $Location
Start-Process -FilePath ".\Add2Exchange_Upgrade.msi" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting