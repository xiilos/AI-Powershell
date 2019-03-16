if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Remove Recovery and Migration Manager

Write-Host "Removing Recovery and Migration Manager"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Recovery and Migration Manager'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary\RMM Sub Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\RMM Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "RMM Upgrades Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\RMM Upgrades"
}

#Downloading Recovery and Migration Manager

Write-Host "Downloading Recovery and Migration Manager"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/RMM-Enterprise/Upgrades/rmm-enterprise.exe"
$Output = "C:\zlibrary\RMM Upgrades\rmm-enterprise.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Recovery and Migration Manager

Write-Host "Unpacking Recovery and Migration Manager"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\RMM Upgrades"
Start-Process "c:\zlibrary\RMM Upgrades\rmm-enterprise.exe" -wait
Write-Host "Done"

#Installing Recovery and Migration Manager

Write-Host "Installing Recovery and Migration Manager"
$Location = Get-ChildItem -Path . -Recurse | Where-Object {$_.LastWriteTime -gt (Get-Date).AddSeconds(-1)}
Push-Location $Location
Start-Process -FilePath "./rmm-enterprise.msi" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 2

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting