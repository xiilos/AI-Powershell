<#
        .SYNOPSIS
        Timed A2E SQL Backup

        .DESCRIPTION
        This is a part of a scheduled task to run and backup A2E SQl DB every 3 days
        5 version retention by default

        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>

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
$BackupDirs = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_DB_Backup.txt" -ErrorAction SilentlyContinue #This is were the default Backup DB Files are stored



# Script #

#Check if Console Open
$Console = Get-Process "Add2Exchange Console" -ErrorAction SilentlyContinue
if ($Console) {
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10003 -EntryType FailureAudit -Message "Add2Exchange Console is Open and cannot backup the A2E SQL Databse $_.Exception.Message"
    Get-PSSession | Remove-PSSession
    Exit

}


#Shutting Down Services
Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 10
Write-Host "Done"
Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service –StartupType Disabled

#Stop The Add2Exchange Agent
Write-Host "Stopping the Agent. Please Wait."
Start-Sleep -s 10
$Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
if ($Agent) {
    Write-Host "Waiting for Agent to Exit"
    Start-Sleep -s 10
    if (!$Agent.HasExited) {
        $Agent | Stop-Process -Force
    }
}
Write-Host "Stopping Add2Exchange SQL Service"
Stop-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 10
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
Copy-Item $currentDB -Destination ($BackupDirs + "\" + (Get-Date -format "dd-MMM-yyyy HH.mm.ss")) -Recurse -ErrorAction SilentlyContinue -ErrorVariable DB1
If ($DB1) {
    Write-Host "Error.....Cannot Find A2E Database."
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Write-Host "Starting Services..."
    Start-Sleep -S 2
    Start-Service -Name "SQL Server (A2ESQLSERVER)"
    Start-Sleep -s 5
    Set-Service -Name "Add2Exchange Service" -StartupType Automatic
    Start-Sleep -s 5
    Start-Service -Name "Add2Exchange Service"
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10004 -EntryType Information -Message "Add2Exchange did not succesfully backup the DB. Starting up services."
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}



Write-Host "Finished Backing up All Files"
Start-Sleep -S 5


Write-Host "Starting Add2Exchange SQL Service"
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 5

Set-Service -Name "Add2Exchange Service" -StartupType Automatic
Start-Sleep -s 5
Start-Service -Name "Add2Exchange Service"


#Write to Event Log
Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10002 -EntryType SuccessAudit -Message "Add2Exchange Succesfully Backed up Database."


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting