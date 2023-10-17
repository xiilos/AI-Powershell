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

#Variables
#$DBInstance = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Softwareù\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue


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
    Install-Module ùName SQLSERVER -Force
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
Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service ùStartupType Disabled

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
Write-Host "Checking SQl Version..."
Push-Location SQLSERVER:\SQL\localhost 
$BuildVersion = Get-ChildItem | Select-Object Version -ExpandProperty Version


#Main Upgrade Menu
Push-Location "C:\zlibrary\SQL Upgrade"

#SQL 2008
if ($BuildVersion.Major -eq '10' -and $BuildVersion.Build -le '5900') {
    Write-Host "Downloading and Upgrading to SQL Express 2008 SP4"
        
}


if ($BuildVersion.Major -eq '10' -and $BuildVersion.Build -ge '5999') {
    Write-Host "Downloading and Upgrading to SQL Express 2012 SP4"
      
}


#----------------------------------------------------------------------


#SQL 2012
if ($BuildVersion.Major -eq '11' -and $BuildVersion.Build -le '7000') {
        
    Write-Host "Downloading and Upgrading to SQL Express 2012 SP4"
}


if ($BuildVersion.Major -eq '11' -and $BuildVersion.Build -ge '7000') {
        
    Write-Host "Exporting Old DB, Installing SQL Express 2022 and Re-importing DB"
}

#----------------------------------------------------------------------

#SQL 2017+
if ($BuildVersion.Major -eq '14') {
        
    Write-Host "Downloading and Upgrading to SQL Express 2022"
}


#----------------------------------------------------------------------






