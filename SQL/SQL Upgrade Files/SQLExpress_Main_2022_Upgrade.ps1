<#
        .SYNOPSIS
        Backup and Store A2E DB
        Find SQL Express Version and upgrade accordingly
        Upgrades SQL Express 8x to SQL Express 2022
        Note* SQL 2008 must be at least SP4 to update to SQL Express 2012
        Note* SQL Express 2012 SP4 last version for x86. Must export and import DB into fresh SQL Express 2022

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

#Warning
$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Caution... Only run this tool on installtions that solely host the Add2Exchange Database locally. Click OK to Continue. or Cancel if you are unsure to quit", 0, "WARNING!!", 0x1)
if ($answer -eq 2) {
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}


#Check DIBSERVER and SERVER Name Match
#Database Server
$DBServer = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBServer" -ErrorAction SilentlyContinue
#Server Name
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server" -ErrorAction SilentlyContinue

If ($DBServer -eq $ServerName) {
    Write-Host "DIB Server Names Match!"
}

Else {

    Write-Host "SQL Server name and localhost Server name do not match. This tool cannot upgrade SQl Express on unmatched server names."
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}




#Check for SQL Module
Write-Host "Checking for SQL Server Module"

IF (Get-Module -ListAvailable -Name SQLSERVER) {
    Write-Host "SQL Server Module Exists"

    $InstalledSQLMod = ((Get-Module -Name SQLSERVER -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()

    $LatestSQLMod = (Find-Module -Name SQLSERVER).Version.ToString()

    [PSCustomObject]@{
        Match = If ($InstalledSQLMod -eq $LatestSQLMod) { Write-Host "You are on the latest Version" } 

        Else {
            Write-Host "Upgrading Modules..."
            Update-Module -Name SQLSERVER -Force
            Write-Host "Success"
        }

    }


} 
Else {
    Write-Host "Module Does Not Exist"
    Write-Host "Downloading Sql Server Module..."
    Install-Module -Name SQLSERVER -Force -AllowClobber
    Write-Host "Success"
}

Import-Module -Name SQLSERVER

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
Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service -StartupType Disabled

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
Write-Host "Done"



#Check SQL Version
Write-Host "Checking SQL Version..."
Push-Location SQLSERVER:\SQL\localhost 
$BuildVersion = Get-ChildItem | Select-Object Version -ExpandProperty Version


#Main Upgrade Menu
Push-Location "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\SQL_Upgrade_Files"



#SQL 2008
if ($BuildVersion.Major -eq '10' -and $BuildVersion.Build -le '5900') {
    Write-Host "SQL Express 2008 currently installed" -ForegroundColor Green
    Write-Host "Downloading and Upgrading to SQL Express 2008 SP4"
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL8x_to_SQL8xSP4 .ps1 -wait
}


if ($BuildVersion.Major -eq '10' -and $BuildVersion.Build -ge '5999') {
    Write-Host "SQL Express 2008 SP4 currently installed" -ForegroundColor Green
    Write-Host "Downloading and Upgrading to SQL Express 2012 SP4"
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL8x_to_SQL12x.ps1 -wait
}


#----------------------------------------------------------------------


#SQL 2012
if ($BuildVersion.Major -eq '11' -and $BuildVersion.Build -le '7000') {
    Write-Host "SQL Express 2012 currently installed" -ForegroundColor Green
    Write-Host "Downloading and Upgrading to SQL Express 2012 SP4"
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL12x_to_SQL12xSP4 .ps1 -wait
}


if ($BuildVersion.Major -eq '11' -and $BuildVersion.Build -ge '7000') {
    Write-Host "SQL Express 2012 SP4 currently installed" -ForegroundColor Green
    Write-Host "Downloading and Upgrading to SQL Express 2022"
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL12x_to_SQL22x.ps1 -wait
}

#----------------------------------------------------------------------

#SQL 2017
if ($BuildVersion.Major -eq '14') {
    Write-Host "Downloading and Upgrading to SQL Express 2022" -ForegroundColor Green
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL17x_to_SQL22x.ps1 -wait
}



#----------------------------------------------------------------------

#SQL 2019
if ($BuildVersion.Major -eq '15') {
    Write-Host "Downloading and Upgrading to SQL Express 2022" -ForegroundColor Green
    Write-Host "Click Enter to continue with upgrade when ready"
    Pause
    Start-Process Powershell .\SQL17x_to_SQL22x.ps1 -wait
}


#----------------------------------------------------------------------

#SQL 2022
if ($BuildVersion.Major -eq '16') {
    Write-Host "You are already on the latest SQL 2022 Express" -ForegroundColor Green
    Pause
}


#----------------------------------------------------------------------



#SQL Management Studio
$wshell2 = New-Object -ComObject Wscript.Shell
$answer2 = $wshell2.Popup("Would you like you install SQL Management Studio? Click OK to start the silent installation, or Cancel to quit", 0, "WARNING!!", 0x1)

if ($answer2 -eq 1) {
    Write-Host "Downloading and silently installing SQL Management Studio"
    Start-Process Powershell .\SQL_Management_Studio_Quiet_Install.ps1
}


if ($answer2 -eq 2) {
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}



Write-Host "Optional; you can now cleanup old SQL Express and SQL Studio installs."
Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting




