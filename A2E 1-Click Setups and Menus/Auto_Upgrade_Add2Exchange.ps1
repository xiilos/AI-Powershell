if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
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
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\Add2Exchange Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Upgrades exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Add2Exchange Upgrades"
}

#Downloading Add2Exchange

Write-Host "Downloading Add2Exchange"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/Upgrades/a2e-enterprise_upgrade.exe"
$Output = "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Exchange

Write-Host "Unpacking Add2exchange"
Write-Host "please Wait....."
Push-Location "c:\zlibrary\Add2Exchange Upgrades"
Start-Process "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe" -wait
Write-Host "Done"

#Installing Add2Exchange

Write-Host "Installing Add2Exchange"
$Location = Get-ChildItem -Path . -Recurse | Where-Object {$_.LastWriteTime -gt (Get-Date).AddSeconds(-1)}
Push-Location $Location
Start-Process -FilePath ".\Add2Exchange_Upgrade.msi" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 3

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting