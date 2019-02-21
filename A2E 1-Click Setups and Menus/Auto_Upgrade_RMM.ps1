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

#Remove Recovery and Migration Manager

Write-Host "Removing Recovery and Migration Manager"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Recovery and Migration Manager'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary\RMM Sub Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\RMM"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "RMM exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path C:\zlibrary\RMM
 }

#Downloading Recovery and Migration Manager

Write-Host "Downloading Recovery and Migration Manager"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/RMM-Enterprise/Upgrades/rmm-enterprise.exe"
$Output = "c:\zlibrary\RMM\rmm-enterprise.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Recovery and Migration Manager

Write-Host "Unpacking Recovery and Migration Manager"
Write-Host "please Wait....."
Set-Location c:\zlibrary\RMM
Start-Process "c:\zlibrary\RMM\rmm-enterprise.exe" -wait
Write-Host "Done"

#Installing Recovery and Migration Manager

Write-Host "Installing Recovery and Migration Manager"
$Location = Get-ChildItem -Path . -Recurse | Where-Object {$_.LastWriteTime -gt (Get-Date).AddSeconds(-1)}
Set-Location $Location
Start-Process -FilePath "./rmm-enterprise.msi" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 3

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting