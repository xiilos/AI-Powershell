if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Test for Upgrade Eligibility

#Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1*" | Select-Object LicenseKeyASMDate, LicenseKeyCSMDate, LicenseKeyNSMDate, LicenseKeyOSMDate, LicenseKeyPSMDate, LicenseKeyTSMDate

$LicenseKeyASMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyASMDate" -ErrorAction SilentlyContinue
$LicenseKeyCSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCSMDate" -ErrorAction SilentlyContinue
$LicenseKeyNSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNSMDate" -ErrorAction SilentlyContinue
$LicenseKeyOSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOSMDate" -ErrorAction SilentlyContinue
$LicenseKeyPSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPSMDate" -ErrorAction SilentlyContinue
$LicenseKeyTSMDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTSMDate" -ErrorAction SilentlyContinue

$Today = Get-Date -format MM/dd/yyy

if ($Today -ge $LicenseKeyASMDate,
    $Today -ge $LicenseKeyCSMDate,
    $Today -ge $LicenseKeyNSMDate,
    $Today -ge $LicenseKeyOSMDate,
    $Today -ge $LicenseKeyPSMDate,
    $Today -ge $LicenseKeyTSMDate)
{
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    $answer = $wshell.Popup("Your Add2Exchange License has Expired! Click OK to continue with the upgrade, or Cancel to Quit.", 0, "ATTENTION!! Add2Exchange Licensing", 0 + 1)
    if ($answer -eq 2) { Break }
}

else {
    Write-Host "We are good to go"
}

Pause


$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop

$answer = $wshell.Popup("If yourCurrent License Dates:
        $LicenseKeyASMDate
        $LicenseKeyCSMDate
        $LicenseKeyNSMDate
        $LicenseKeyOSMDate
        $LicenseKeyPSMDate
        $LicenseKeyTSMDate", 0, "ATTENTION!! Add2Exchange Licensing", 0 + 1)
if ($answer -eq 2) { Break }





# Test for FTP

$FTP = Test-NetConnection -ComputerName ftp.diditbetter.com -Port 21
if ( $(Try { Test-Path $FTP.trim() } Catch { $false }) ) {

    Write-Host "No FTP Access... Taking you to Downloads...."
    Invoke-Item "http://support.diditbetter.com/Secure/Login.aspx?returnurl=/downloads.aspx"
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#Stop Add2Exchange Service

Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 2
Write-Host "Done"

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


#Remove Add2Exchange

Write-Host "Removing Add2Exchange"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\Add2Exchange Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Upgrades Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Add2Exchange Upgrades"
}

#Downloading Add2Exchange

Write-Host "Downloading Add2Exchange"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/Upgrades/a2e-enterprise_upgrade.exe"
$Output = "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Exchange

Write-Host "Unpacking Add2exchange"
Write-Host "please Wait....."
Push-Location "c:\zlibrary\Add2Exchange Upgrades"
Start-Process "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe" -wait
Start-Sleep -Seconds 2
Write-Host "Done"

#Installing Add2Exchange
Do {
    Write-Host "Installing Add2Exchange"
    $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Push-Location $Location
    Start-Process -FilePath ".\Add2Exchange_Upgrade.msi" -wait -ErrorAction Inquire -ErrorVariable InstallError;
    Write-Host "Finished...Upgrade Complete"

    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))

#Setting the Service to Delayed Start
Write-Host "Setting up Add2Exchange Service to Delayed Start"
sc.exe config "Add2Exchange Service" start= delayed-auto
Write-Host "Done"

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting