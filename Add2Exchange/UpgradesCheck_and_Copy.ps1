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

    Write-Host "Latest Add2Exchange Upgrade Already Exists in Upgrades Folder" -ForegroundColor Green
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades"
    Remove-Item -path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades\a2e-enterprise_upgrade.exe" -ErrorAction SilentlyContinue
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades\$Latest" -NewName "a2e-enterprise_upgrade.exe" -ErrorAction SilentlyContinue
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\Upgrades"
    Write-Host "Done"
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Grab the Latest A2E Full Install Build

$Dir = "\\diditbetter\DFS\FTP\download\A2E-Enterprise"
$Filter = "a2e-enterprise.*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Installs

Write-Host "Checking New Installs Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest Add2Exchange Enterprise Full Already Exists in New Installs Folder" -ForegroundColor Green
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs"
    Remove-Item -path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs\a2e-enterprise.exe" -ErrorAction SilentlyContinue
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs\$Latest" -NewName "a2e-enterprise.exe" -ErrorAction SilentlyContinue
    Copy-Item "\\diditbetter\DFS\FTP\download\A2E-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\New Installs"
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

    Write-Host "Latest Add2Outlook Upgrade Already Exists in Upgrades Folder" -ForegroundColor Green
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\Add2Outlook\$Latest" -Destination "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades"
    Remove-Item -path "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades\Add2Outlook Full Installation.EXE" -ErrorAction SilentlyContinue
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades\$Latest" -NewName "Add2Outlook Full Installation.EXE" -ErrorAction SilentlyContinue
    Copy-Item "\\diditbetter\DFS\FTP\download\Add2Outlook\$Latest" -Destination "\\diditbetter\DFS\FTP\download\Add2Outlook\Upgrades"
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

    Write-Host "Latest RMM Upgrade Already Exists in Upgrades Folder" -ForegroundColor Green
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\RMM-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades"
    Remove-Item -path "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades\rmm-enterprise.exe" -ErrorAction SilentlyContinue
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades\$Latest" -NewName "rmm-enterprise.exe" -ErrorAction SilentlyContinue
    Copy-Item "\\diditbetter\DFS\FTP\download\RMM-Enterprise\$Latest" -Destination "\\diditbetter\DFS\FTP\download\RMM-Enterprise\Upgrades"
    Write-Host "Done"
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Grab the Latest ToolKit Upgrade Build

$Dir = "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit"
$Filter = "Add2Outlook ToolKit Full Installation*"
$Latest = Get-ChildItem -Path $Dir -Filter $Filter | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$Latest.Name

#Testing to See if File Already Exists in Upgrades

Write-Host "Checking Upgrades Folder for Existing Installs"
$TestPath = "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\Upgrades\$Latest"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Latest ToolKit Upgrade Already Exists in Upgrades Folder" -ForegroundColor Green
}
Else {
    Write-Host "Copying Files Needed..."
    Copy-Item "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\$Latest" -Destination "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\Upgrades"
    Remove-Item -path "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\Upgrades\Add2Outlook ToolKit Full Installation.exe" -ErrorAction SilentlyContinue
    Rename-Item -Path "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\Upgrades\$Latest" -NewName "Add2Outlook ToolKit Full Installation.exe" -ErrorAction SilentlyContinue
    Copy-Item "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\$Latest" -Destination "\\diditbetter\DFS\FTP\download\Add2Outlook Toolkit\Upgrades"
    Write-Host "Done"
}

Write-Host "All Upgrades Have Been Copied Over!!!" -ForegroundColor Green

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Copy Tools to A2O

Write-Host "Copying Setup files to Add2Outlook " -ForegroundColor Green

Copy-Item "\\diditbetter\dev\DevLimAccess\a2e-enterprise\Contrib\Setup.zip" -Destination "\\diditbetter\dev\DevLimAccess\Add2Outlook\contrib\Setup.zip" -Force -verbose
Copy-Item "\\diditbetter\dev\DevLimAccess\a2e-enterprise\Contrib\Setup.zip" -Destination "\\diditbetter\DFS\FTP\download\A2E-Enterprise\A2E Tools\Setup.zip" -Force -verbose


Write-Host "All files Copied Over!!!" -ForegroundColor Green

Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting