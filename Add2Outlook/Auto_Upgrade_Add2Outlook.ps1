if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Stop Menu Process
Stop-Process -Name "DidItBetterSupportMenu" -Force -ErrorAction SilentlyContinue

# Test for FTP

try {
    $FTP = New-Object System.Net.Sockets.TcpClient("ftp.diditbetter.com", 21)
    $FTP.Close()
    Write-Host "Connectivity OK."
}
catch {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    $wshell.Popup("No FTP Access... Taking you to Downloads.... Click OK or Cancel to Quit.", 0, "ATTENTION!!", 0 + 1)
    Start-Process "http://support.diditbetter.com/Secure/Login.aspx?returnurl=/downloads.aspx"
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#Create zLibrary\Add2Outlook Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\Add2Outlook"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Outlook Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Add2Outlook"
}

#Downloading Add2Outlook

Write-Host "Downloading Add2Outlook"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/Add2Outlook/Upgrades/Add2Outlook%20Full%20Installation.EXE"
$Output = "C:\zlibrary\Add2Outlook\Add2OutlookToolKitFullInstallation.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Outlook

Write-Host "Unpacking Add2Outlook"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\Add2Outlook"
Write-Host "Done"

#Installing Add2Outlook

Write-Host "Installing Add2Outlook"
Start-Process -FilePath "./Add2OutlookToolKitFullInstallation.exe" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 2

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting