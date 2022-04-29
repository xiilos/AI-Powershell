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
#$BackupDirs = $Install + 'Database Backup\' #This is were the default Backup DB Files are stored



#Script#

#Folder path for Backups
Write-Host "Pick a destination for your backups"
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
[void]$FolderBrowser.ShowDialog()
$BackupDirs = $FolderBrowser.SelectedPath #This is were the chosen Backup DB Files are stored

<#
#Check and Create Default DB Backup Locations
$TestPath1 = "$Install\Database Backup"
if ( $(Try { Test-Path $TestPath1.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Backup Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "$Install\Database Backup"
}
#>

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
Copy-Item $currentDB -Destination ($BackupDirs + "\" + (Get-Date -format "dd-MMM-yyyy HH.mm.ss")) -Recurse -ErrorAction SilentlyContinue -ErrorVariable DB1
If ($DB1) {
    Write-Host "Error.....Cannot Find A2E Database."
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Pause
}



Write-Host "Finished Backing up All Files"
Start-Sleep -S 5


Write-Host "Starting Add2Exchange SQL Service"
Start-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 5

Set-Service -Name "Add2Exchange Service" -StartupType Automatic

#Write to Event Log
Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10002 -EntryType SuccessAudit -Message "Add2Exchange Succesfully Backed up Database."


#Creating the Task
$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Would you like to make this into an Automated Task?", 0, "Backup Complete", 0x1)
if ($answer -eq 1) { 
    #Checking to see if task is already in use
    if (Get-ScheduledTask "Add2Exchange Database Backup" -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName "Add2Exchange Database Backup" -Confirm:$false
    }

    Write-Host "Pick a destination for your backups"
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()
    $BackupDirs = $FolderBrowser.SelectedPath #This is were the chosen Backup DB Files are stored
    Write-Host "Creating A2E Backup Task"
    $Backupdirs | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_DB_Backup.txt"


    $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
    Set-Location $Location
    $Repeater = (New-TimeSpan -days 3)
    $Duration = ([timeSpan]::maxvalue)
    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed_A2E_SQL_Backup.ps1"'
    $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange SQL Backup" -Description "Backs up the A2E SQL Database every once a week" -User $UserID -Password $Password
    Write-Host "Done"

}

if ($answeer -eq 2) {
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit 
}



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