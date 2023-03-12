<#
        .SYNOPSIS
        SQL 2008 Upgrade

        .DESCRIPTION
        Upgrades SQL 2008 A2E DB to SQL 2012
        Can only upgrade to 2012, since that was the last 32bit SQL avialable

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
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

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
    $wshell.Popup("Connection failed with error: $($_.Exception.Message)... Taking you to Downloads.... Click OK or Cancel to Quit.", 0, "ATTENTION!!", 0 + 1)
    Start-Process "http://support.diditbetter.com/Secure/Login.aspx?returnurl=/downloads.aspx"
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#Downloading SQL Express 2012
Write-Host "Downloading SQL Express 2012"
Write-Host "Please Wait......"

$URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQLEXPR_x86_ENU.exe"
$Output = "C:\zlibrary\SQL Upgrade\SQLEXPR_x86_ENU.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Launching SQL Express
Write-Host "Starting SQL Express 2012 Upgrade"
Write-Host "Please Wait....."
Push-Location "C:\zlibrary\SQL Upgrade"
Start-Process -FilePath "./SQLEXPR_x86_ENU.exe" -wait -ErrorAction Stop


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting