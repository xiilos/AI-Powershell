if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Grab the Latest A2E Upgrade Build

$Dir = "\\diditbetter\DFS\FTP\download\A2E-Enterprise"
$Filter = "a2e-enterprise_upgrade*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Upgrades

Write-Host "Checking Upgrades Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest Add2Exchange Upgrade Already Exists in Upgrades Folder"
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades"
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades\$Latest" -NewName "a2e-enterprise_upgrade.exe" -ErrorAction SilentlyContinue
    Write-Host "Done"
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Grab the Latest A2E Full Install Build

$Dir = "\\diditbetter\DFS\FTP\download\A2E-Enterprise"
$Filter = "a2e-enterprise.*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Upgrades

Write-Host "Checking New Installs Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest Add2Exchange Enterprise Full Already Exists in New Installs Folder"
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs"
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs\$Latest" -NewName "a2e-enterprise.exe" -ErrorAction SilentlyContinue
    Write-Host "Done"
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Grab the Latest A2O Upgrade Build

$Dir = "\\diditbetter\DFS\FTP\download\Add2Outlook"
$Filter = "Add2Outlook Full Installation*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Upgrades

Write-Host "Checking Upgrades Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest Add2Outlook Upgrade Already Exists in Upgrades Folder"
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\Add2Outlook\$Latest" -Destination "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades"
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades\$Latest" -NewName "Add2Outlook Full Installation.EXE" -ErrorAction SilentlyContinue
    Write-Host "Done"
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Grab the Latest RMM Upgrade Build

$Dir = "\\diditbetter\DFS\FTP\download\RMM-Enterprise"
$Filter = "rmm-enterprise*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Upgrades

Write-Host "Checking Upgrades Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest RMM Upgrade Already Exists in Upgrades Folder"
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\RMM-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades"
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades\$Latest" -NewName "rmm-enterprise.exe" -ErrorAction SilentlyContinue
    Write-Host "Done"
}

Write-Host "All Upgrades Have Been Copied Over!!!"

Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting