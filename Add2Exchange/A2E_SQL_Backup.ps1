if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Variables
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue #Current Add2Exchange Installation Path
$Staging = $Install + 'Database\Temp' #Temporary Staging area for SQL backup files
$BackupDirs = $Install + 'Database\A2E_SQL_Backup'
$Date = (Get-Date -format yyyy-MM-dd)


# Script #

#Check and Create DB Backup Locations

$TestPath1 = "$Install\Database\A2E_SQL_Backup"
if ( $(Try { Test-Path $TestPath1.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Backup Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "$Install\Database\A2E_SQL_Backup"
}

$TestPath2 = "$Install\Database\Temp"
if ( $(Try { Test-Path $TestPath2.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Temp Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "$Install\Database\Temp"
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
Write-Host "Backinp Up Add2Exchange SQL Files"

Copy-Item "$Install\Database\A2E.mdf" -Destination "$Staging" -Recurse -ErrorAction SilentlyContinue -ErrorVariable DB1
If ($DB1) {
    Write-Host "Error.....Cannot Find A2E.mdf Database."
    Pause
}
Copy-Item "$Install\Database\A2E_log.LDF" -Destination "$Staging" -Recurse -ErrorAction SilentlyContinue -ErrorVariable DB2
If ($DB2) {
    Write-Host "Error.....Cannot Find A2E_Log.LDF Database."
    Pause
}

Write-Host "Starting Add2Exchange SQL Service"
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 5

Write-Host "Creating A2E SQL Archive $Date"

Compress-Archive -LiteralPath "$Staging\A2E.mdf","$Staging\A2E_log.LDF" -DestinationPath ($BackupDirs+"\" + (Get-Date -format yyyy-MM-dd) +"-"+(Get-Random -Maximum 900000) + ".zip") -CompressionLevel Optimal -Force

Start-Sleep -s 5

Get-Childitem -Path $Staging -Recurse -Force | Remove-Item -Confirm:$false -Recurse

Write-Host "All Add2Exchange SQL Backups are stored in $Backupdirs"
Write-Host "Backup Complete..."
Pause


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting