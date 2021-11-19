if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Variables
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue #Current Add2Exchange Installation Path
$CurrentDB = $Install + 'Database\' #Current Database Location
$BackupDirs = $Install + 'Database Backup\' #This is were the Backup DB Files are stored



# Script #

#Check and Create DB Backup Locations
$TestPath1 = "$Install\Database Backup"
if ( $(Try { Test-Path $TestPath1.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Backup Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "$Install\Database Backup"
}


#Shutting Down Services
Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 2
Write-Host "Done"
Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service –StartupType Disabled

#Stop The Add2Exchange Agent
Write-Host "Stopping the Agent. Please Wait."
Start-Sleep -s 5
$Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
if ($Agent) {
    Write-Host "Waiting for Agent to Exit"
    Start-Sleep -s 5
    if (!$Agent.HasExited) {
        $Agent | Stop-Process -Force
    }
}
Write-Host "Stopping Add2Exchange SQL Service"
Stop-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 5
Write-Host "Done"


#Backing Up SQL Files
Write-Host "Backing Up Add2Exchange SQL Files"


#Backup Retention (5)
$Path = $BackupDirs
Push-Location $Path

$FolderCount = (Get-ChildItem -Path $Path | where-object { $_.PSIsContainer }).Count
If ($FolderCount -gt 5) {
    $Object = Get-Childitem -Path $Path | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -Last 1
    Write-Host "Removing $Object" -ForegroundColor Red
    Remove-Item $Object -Recurse -Force
    Write-Host "Done"
}

  
#Back Up Files
Copy-Item $currentDB -Destination ($BackupDirs + (Get-Date -format "dd-MMM-yyyy HH.mm.ss")) -Recurse -ErrorAction SilentlyContinue -ErrorVariable DB1
If ($DB1) {
    Write-Host "Error.....Cannot Find A2E Database."
    Pause
}



Write-Host "Finished Backing up All Files"
Start-Sleep -S 5


Write-Host "Starting Add2Exchange SQL Service"
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 5


$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("All Add2Exchange SQL Backups are stored in $Backupdirs", 0, "Backup Complete", 0x1)
if ($answer -eq 2) { 
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit 
}

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting