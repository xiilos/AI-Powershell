<#
        .SYNOPSIS
PowerShell Script to Upgrade SQL Server Express 2012 SP4 to 2022

Ensure you run PowerShell as Administrator

Ensure to adjust paths and instance names as per your environment.

1. Backup Databases
Implement backup logic as per your environment & requirement.

2. Verify SQL Server 2008 SP4 is installed
Verify manually or add script logic as per your requirement.

3. Install SQL Express 2022

        .DESCRIPTION
      Will upgrade SQL Express 2012 SP4+ to SQL Express 2022

        .NOTES
        Version:        1.3
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
$installerPath = "C:\zlibrary\SQL Upgrade\SQL2022-SSEI-Expr.exe"



    
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
$backupFolder = "C:\zlibrary\SQL Backup\$((Get-Date).ToString('yyyy-MM-dd-hh-mm-ss'))"

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
    } else {
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
    Write-Host "Downloading SQL Express 2022"
    $URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SQL2022-SSEI-Expr.exe"
    $Output = "C:\zlibrary\SQL Upgrade\SQL2022-SSEI-Expr.exe"
    $Start_Time = Get-Date
    
    (New-Object System.Net.WebClient).DownloadFile($URL, $Output)
    
    Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"
    
#Download SQL Express Config.ini
$URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/Microsoft_SQL_Server_Express_2022.ini"
$Output = "C:\zlibrary\SQL Upgrade\Microsoft_SQL_Server_Express_2022.ini"
$Start_Time = Get-Date
    
    (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

    Write-Host "Finished Downloading"
    }
    catch {
        Write-Host "Donwloading failed through Powershell, trying another way...." -ForegroundColor Red
        Start-Process "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SQL2022-SSEI-Expr.exe"
        Start-Process "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/Microsoft_SQL_Server_Express_2022.ini"
        Write-Host "Ensure to put both the downloaded SQL Express installer and .ini file in the folder C:\zlibrary\SQL Upgrade\* When finished, click Enter to continue"
        Pause
    }
    

#Remove SQL Express 2012
# Path to the SQL Server 2012 Express setup executable
Write-Host "Removing SQL Express 2012. Please Wait..."
$SetupPath = "C:\Program Files (x86)\Microsoft SQL Server\110\Setup Bootstrap\SQLServer2012\setup.exe"
$instanceName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue

# Run the uninstall command silently
Start-Process -FilePath $SetupPath -ArgumentList "/Action=Uninstall /FEATURES=SQLEngine /INSTANCENAME=$InstanceName /Q" -Wait
Write-Output "SQL Express 2012 is now Uninstalled"


#Rename MDF and LDF files
$dirPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Database"

# Check if the directory exists
if (Test-Path $dirPath) {
    # Define the old and new file names
    $oldMdf = Join-Path $dirPath "A2E.MDF"
    $newMdf = Join-Path $dirPath "Original_A2E.MDF"
    $oldLdf = Join-Path $dirPath "A2E_Log.LDF"
    $newLdf = Join-Path $dirPath "Original_A2E_Log.LDF"

    # Rename A2E.MDF to Original_A2E.MDF
    if (Test-Path $oldMdf) {
        Rename-Item -Path $oldMdf -NewName $newMdf
        Write-Host "Renamed $oldMdf to $newMdf"
    }
    else {
        Write-Host "$oldMdf not found!"
    }

    # Rename A2E_Log.LDF to Original_A2E_Log.LDF
    if (Test-Path $oldLdf) {
        Rename-Item -Path $oldLdf -NewName $newLdf
        Write-Host "Renamed $oldLdf to $newLdf"
    }
    else {
        Write-Host "$oldLdf not found!"
    }

}
else {
    Write-Host "Directory $dirPath not found!"
}


#Backup and Remove A2E SQL Reg Key
REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" C:\zlibrary\Add2Exchange.Reg
Remove-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\OpenDoor Software®\Add2Exchange" -Name "SQLTables"


#Launch the installer with arguments
Write-Host "Installing SQL Express 2022..."
Push-Location "C:\zLibrary\SQL Upgrade"
Start-Process -FilePath "$installerPath" -ArgumentList "/ConfigurationFile=Microsoft_SQL_Server_Express_2022.ini /Q /IACCEPTSQLSERVERLICENSETERMS" -Wait -PassThru
Write-Host "Finished...Installation Complete.."
Write-Host "Press Enter to Proceed with upgrade"
    
If ($InstallError) { 
    Write-Warning -Message "Something Went Wrong with the Installation!"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}
    
pause

#Starting the Add2Exchange Console
Write-Host "Starting the Add2Exchange Console"
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
Push-Location $Install
Start-Process "./Console/Add2Exchange Console.exe"
Write-Host "Once the A2E SQL tables are created, Choose the OK Button, wait for the Add2exchange Console to load and then hit enter for Powershell to continue"
Pause

#Close the Console
Stop-Process -Name "Add2Exchange Console"
Start-Sleep -s 5

#Replacing SQL Files
Write-Host "Re-Importing Database and restarting the Add2Exchange Console"


Write-Host "Stopping Add2Exchange SQL Service"
Stop-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
Write-Host "Done"

#Rename New Created MDF and LDF files to New_
$dirPathnew = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Database"
# Check if the directory exists
if (Test-Path $dirPathnew) {
    # Define the old and new file names
    $zoldMdf = Join-Path $dirPathnew "A2E.MDF"
    $znewMdf = Join-Path $dirPathnew "New_A2E.MDF"
    $zoldLdf = Join-Path $dirPathnew "A2E_Log.LDF"
    $znewLdf = Join-Path $dirPathnew "New_A2E_Log.LDF"

    # Rename A2E.MDF to zA2E.MDF
    if (Test-Path $zoldMdf) {
        Rename-Item -Path $zoldMdf -NewName $znewMdf
        Write-Host "Renamed $zoldMdf to $znewMdf"
    }
    else {
        Write-Host "$zoldMdf not found!"
    }

    # Rename A2E_Log.LDF to zA2E_Log.LDF
    if (Test-Path $zoldLdf) {
        Rename-Item -Path $zoldLdf -NewName $znewLdf
        Write-Host "Renamed $zoldLdf to $znewLdf"
    }
    else {
        Write-Host "$zoldLdf not found!"
    }

}
else {
    Write-Host "Directory $dirPathnew not found!"
}


#Rename Original MDF and LDF files to Original Names
$dirPathnew = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Database"
# Check if the directory exists
if (Test-Path $dirPathnew) {
    # Define the old and new file names
    $zoldMdf = Join-Path $dirPathnew "Original_A2E.MDF"
    $znewMdf = Join-Path $dirPathnew "A2E.MDF"
    $zoldLdf = Join-Path $dirPathnew "Original_A2E_Log.LDF"
    $znewLdf = Join-Path $dirPathnew "A2E_Log.LDF"

    # Rename A2E.MDF to zA2E.MDF
    if (Test-Path $zoldMdf) {
        Rename-Item -Path $zoldMdf -NewName $znewMdf
        Write-Host "Renamed $zoldMdf to $znewMdf"
    }
    else {
        Write-Host "$zoldMdf not found!"
    }

    # Rename A2E_Log.LDF to zA2E_Log.LDF
    if (Test-Path $zoldLdf) {
        Rename-Item -Path $zoldLdf -NewName $znewLdf
        Write-Host "Renamed $zoldLdf to $znewLdf"
    }
    else {
        Write-Host "$zoldLdf not found!"
    }

}
else {
    Write-Host "Directory $dirPathnew not found!"
}

Write-Host "Starting Add2Exchange SQL Service"
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
Write-Host "Done"

Write-Host "Starting The Add2Exchange Console"
Write-Host "Starting the Add2Exchange Console"
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
Push-Location $Install
Start-Process "./Console/Add2Exchange Console.exe"
Write-Host "Installation and Upgrade are Complete."
Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
    
# End Scripting




