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


#Test for FTP

try {
    $FTP = New-Object System.Net.Sockets.TcpClient("ftp.diditbetter.com", 21)
    $FTP.Close()
    Write-Host "Connectivity OK."
}
catch {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    $wshell.Popup("No FTP Access... Taking you to Downloads.... Click OK or Cancel to Quit.", 0, "ATTENTION!!", 0 + 1)
    Start-Process "https://download.microsoft.com/download/F/6/7/F673709C-D371-4A64-8BF9-C1DD73F60990/ENU/x86/SQLEXPR_x86_ENU.exe"
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#Downloading SQL Express 2012

Write-Host "Downloading SQL Express 2012"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/SQL/SQL2012Management/SQLEXPR_x86_ENU.exe"
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