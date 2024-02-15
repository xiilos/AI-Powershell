<#
        .SYNOPSIS
PowerShell Script to Upgrade SQL Server Express 2008 to 2008 SP4

Ensure you run PowerShell as Administrator

Ensure to adjust paths and instance names as per your environment.

1. Backup Databases
Implement backup logic as per your environment & requirement.

2. Verify SQL Server 2008 is installed
Verify manually or add script logic as per your requirement.

3. Install SQL Server 2008 Express SP4

        .DESCRIPTION
      Will upgrade SQL Express 2008 to SQL Express 2008 SP4

        .NOTES
        Version:        1.2
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
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server" -ErrorAction SilentlyContinue
$instanceName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue
$DBname = "A2E"
$installerPath = "C:\zlibrary\SQL Upgrade\SQLServer2008SP4-KB2979596-x86-ENU.exe"
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


#Back up existing SQL Instance
#Define source and destination paths
$instanceLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
$CurrentDB = $instanceLocation + 'Database\' #Current Database Location
$sourcePath = $CurrentDB
$backupFolder = 'C:\zlibrary\SQL Backup'


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

#Copy files to to SQL backup folder
Write-Host "Backing up current SQL Database before upgrade"
# Define the source folder where the files are located
$sourceFolder = $sourcepath

# Define the destination folder where the files will be copied to
$destinationFolder = $backupFolder

# Files to be copied
$filesToCopy = @("a2e.mdf", "a2e_log.ldf")

# Loop through each file in the list and copy them to the destination folder
foreach ($file in $filesToCopy) {
    $sourcePath2 = Join-Path -Path $sourceFolder -ChildPath $file
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $file

    # Check if the source file exists before attempting to copy
    if (Test-Path -Path $sourcePath2) {
        Copy-Item -Path $sourcePath2 -Destination $destinationPath -Force
        Write-Output "Copied: $sourcePath2 to $destinationPath"
        Write-Output "Backup completed to $backupFolder"
    }
    else {
        Write-Output "File not found: $sourcePath2"
        $wshell = New-Object -ComObject Wscript.Shell
        $answer = $wshell.Popup("There seems to be an issue copying the A2E SQL databases. Ensure that your Databases are located in $Sourcepath Click OK to continue or Cancel to Quit", 0, "SQL Backup", 0x1)
        if ($answer -eq 2) { 
            Write-Host "ttyl"
            Get-PSSession | Remove-PSSession
            Exit
        }
    }
}


#Restart SQL Services
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
Write-Host "SQL Services Restarted"

#Test for HTTPS Access
Write-Host "Testing for HTTPS Connectivity"

try {
    $wresponse = Invoke-WebRequest -Uri https://s3.amazonaws.com/dl.diditbetter.com -UseBasicParsing
    if ($wresponse.StatusCode -eq 200) {
        Write-Output "Connection successful"
    }
    else {
        Write-Output "HTTPS Connection failed with status code $($wresponse.StatusCode) Will try a different download method."
    }
}
catch {
#
}

#Download SQL Express Installer
try {
    Write-Host "Downloading SQL Express 2008 SP4"
    $URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SQLServer2008SP4-KB2979596-x86-ENU.exe"
    $Output = "C:\zlibrary\SQL Upgrade\SQLServer2008SP4-KB2979596-x86-ENU.exe"
    $Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

    Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

    Write-Host "Finished Downloading"
}
catch {
    Write-Host "Donwloading failed through Powershell, trying another way...." -ForegroundColor Red
    Start-Process "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SQLServer2008SP4-KB2979596-x86-ENU.exe"
    Write-Host "Ensure to put the downloaded SQL Express installer in the folder C:\zlibrary\SQL Upgrade\* When finished, click Enter to continue"
    Pause
}



#Launch the installer with arguments
Write-Host "Upgrading SQL Express 2008 to SQL Express 2008 SP4..."
Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -ErrorAction Inquire -ErrorVariable InstallError;
Write-Host "Finished...Upgrade Complete.."

If ($InstallError) { 
    Write-Warning -Message "Something Went Wrong with the Upgrade!"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}

Write-Host "Please Reboot to Complete the Upgrade"
Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting