<#
        .SYNOPSIS
PowerShell Script to Upgrade SQL Server Express 2017+ to 2022

Ensure you run PowerShell as Administrator

Ensure to adjust paths and instance names as per your environment.

1. Backup Databases
Implement backup logic as per your environment & requirement.

2. Verify SQL Server 2017+ is installed
Verify manually or add script logic as per your requirement.

3. Install SQL Server 2022 Express

        .DESCRIPTION
      Will upgrade SQL Express 2017+ to SQL Express 2022

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    #Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
    
#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
    
#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append
    
# Script #
    
#Variables
$installerPath = "C:\zlibrary\SQL Upgrade\SQLEXPR_x64_ENU_2022.exe"
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server" -ErrorAction SilentlyContinue
$instanceName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue
$DBname = "A2E"
$arguments = "/Action=Upgrade /Q /INSTANCENAME=$instanceName /IACCEPTSQLSERVERLICENSETERMS"
  

    
#Create zLibrary\A2E SQL Upgrade Directory
Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\SQL Upgrade"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    
    Write-Host "SQL Upgrade Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\SQL Upgrade"
}
    
#Test for HTTPS Access
Write-Host "Testing for HTTPS Connectivity"
    
try {
    $wresponse = Invoke-WebRequest -Uri https://s3.amazonaws.com/dl.diditbetter.com -UseBasicParsing
    if ($wresponse.StatusCode -eq 200) {
        Write-Output "Connection successful"
    }
    else {
        Write-Output "Connection failed with status code $($wresponse.StatusCode)"
    }
}
catch {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    $wshell.Popup("Connection failed with error: $($_.Exception.Message)...Click OK or Cancel to Quit.", 0, "ATTENTION!!", 0 + 1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}
    
#Back up existing SQL Instance
#Define source and destination paths
$instanceLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
$CurrentDB = $instanceLocation + 'Database\*' #Current Database Location
$sourcePath = $CurrentDB
$backupFolder = 'C:\zlibrary\SQL Backup'
$backupFileName = "Backup_$(Get-Date -Format 'yyyyMMddHHmmss').zip"
$backupFilePath = Join-Path -Path $backupFolder -ChildPath $backupFileName
    
# Ensure backup folder exists
if (-not (Test-Path $backupFolder)) {
    New-Item -Path $backupFolder -ItemType Directory
}

#Trim SQL Log
Write-Host "Trimming the SQL transaction Log"

try {
    Invoke-Sqlcmd -ServerInstance "$Servername\$instancename" -Database "$DBname" -Trustservercertificate -Query "DBCC SHRINKFILE('A2E_log', 1);"
}
catch {
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10020 -EntryType FailureAudit -Message "SQL Transaction Log Trim failure $_.Exception.Message"
    Write-Host "Please see the Add2Exchange Log for Errors"
    Pause
    Get-PSSession | Remove-PSSession
    Exit
}

Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10021 -EntryType FailureAudit -Message "Add2Exchange SQL Transaction Log Trimmed Succesfully"


#Stop A2E SQL Service
Write-Host "Stopping Add2Exchange SQL Service"
Stop-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
Write-Host "Done"

#Compress files to a .zip archive
Write-Host "Backing up current SQL Database before upgrade"
Compress-Archive -Path $sourcePath -DestinationPath $backupFilePath
Write-Output "Backup completed to $backupFilePath"
    
#Restart SQL Services
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
Write-Host "SQL Services Restarted"
    
    
#Download SQL Express Installer
Write-Host "Downloading SQL Express 2022"
$URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SQLEXPR_x64_ENU_2022.exe"
$Output = "C:\zlibrary\SQL Upgrade\SQLEXPR_x64_ENU_2022.exe"
$Start_Time = Get-Date
    
    (New-Object System.Net.WebClient).DownloadFile($URL, $Output)
    
Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"
    
Write-Host "Finished Downloading"
    
    
#Launch the installer with arguments 2017+ to 2022
Write-Host "Upgrading SQL Express 2017+ to SQL Express 2022..."
Push-Location "C:\zLibrary\SQL Upgrade"
Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
Write-Host "Finished...Upgrade Complete.."
    
If ($InstallError) { 
    Write-Warning -Message "Something Went Wrong with the Upgrade!"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}
Write-Host "Please Reboot to Complete the Upgrade... Press Enter to Exit"
Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
    
# End Scripting