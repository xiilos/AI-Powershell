if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Test for Upgrade Eligibility
$LicenseKeyDExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyDExpiry" -ErrorAction SilentlyContinue

$Today = Get-Date

#License Varify

#Recovery and Migration Manager--------------

$RMM = if ($LicenseKeyDExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyDExpiry -and $LicenseKeyDExpiry -notlike "") {
    "$LicenseKeyAExpiry !!EXPIRED!!" 
}

Else {
    "$LicenseKeyDExpiry"
}


$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$answer = $wshell.Popup("Please Review your Recovery and Migration Manager License Expiration Dates below.

If your Free Upgrade Period (Software Assurance) has passed and they are expired, you will need to renew the software to continue using it.
If you have already renewed or in that process, please Press Continue. 

Your Software Mainenance Free Upgrade Periods

$RMM for Recovery and Migration Manager

$Today  is today’s date, and any keys for your active functioning modules should expire AFTER this date.

Click OK to continue with your upgrade, or Cancel to Quit.

NOTE* Upgrading Recovery and Migration Manager with expired keys and outside Software Assurance will stop synchronization until a renewal is purchased and new keys are issued.
Also, if the service account can receive email, after you purchase, the keys will automatically be applied, usually without intervention.

", 0, "ATTENTION!! RMM Licensing", 0 + 1)
if ($answer -eq 2) { Break }



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


#Remove Recovery and Migration Manager

Write-Host "Removing Recovery and Migration Manager"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Recovery and Migration Manager'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary\RMM Sub Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\RMM Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "RMM Upgrades Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\RMM Upgrades"
}

#Downloading Recovery and Migration Manager

Write-Host "Downloading Recovery and Migration Manager"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/RMM-Enterprise/Upgrades/rmm-enterprise.exe"
$Output = "C:\zlibrary\RMM Upgrades\rmm-enterprise.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Recovery and Migration Manager

Write-Host "Unpacking Recovery and Migration Manager"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\RMM Upgrades"
Start-Process "c:\zlibrary\RMM Upgrades\rmm-enterprise.exe" -wait
Start-Sleep -Seconds 2
Write-Host "Done"

#Installing Recovery and Migration Manager
Do {
    Write-Host "Installing Recovery and Migration Manager"
    $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Push-Location $Location
    Start-Process -FilePath "./rmm-enterprise.msi" -wait -ErrorAction Inquire -ErrorVariable InstallError;
    Write-Host "Finished...Upgrade Complete"
    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting