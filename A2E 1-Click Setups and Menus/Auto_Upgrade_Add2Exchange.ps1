if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Test for Upgrade Eligibility
$LicenseKeyAExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyAExpiry" -ErrorAction SilentlyContinue
$LicenseKeyCExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCExpiry" -ErrorAction SilentlyContinue
$LicenseKeyEExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyEExpiry" -ErrorAction SilentlyContinue
$LicenseKeyGExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyGExpiry" -ErrorAction SilentlyContinue
$LicenseKeyMExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyMExpiry" -ErrorAction SilentlyContinue
$LicenseKeyNExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNExpiry" -ErrorAction SilentlyContinue
$LicenseKeyOExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOExpiry" -ErrorAction SilentlyContinue
$LicenseKeyPExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPExpiry" -ErrorAction SilentlyContinue
$LicenseKeyTExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTExpiry" -ErrorAction SilentlyContinue

$Today = Get-Date

#License Varify

#Calendars--------------

$Calendars = if ($LicenseKeyAExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyAExpiry -and $LicenseKeyAExpiry -notlike "") {
    "$LicenseKeyAExpiry !!EXPIRED!!" 
}

Else {
    "$LicenseKeyAExpiry"
}

#Contacts--------------

$Contacts = if ($LicenseKeyCExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyCExpiry -and $LicenseKeyCExpiry -notlike "") {
    "$LicenseKeyCExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyCExpiry"
}

#RGM--------------

$RGM = if ($LicenseKeyEExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyEExpiry -and $LicenseKeyEExpiry -notlike "") {
    "$LicenseKeyEExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyEExpiry"
}

#GAL--------------

$GAL = if ($LicenseKeyGExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyGExpiry -and $LicenseKeyGExpiry -notlike "") {
    "$LicenseKeyGExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyGExpiry"
}


#MAIL--------------

$Mail = if ($LicenseKeyMExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyMExpiry -and $LicenseKeyMExpiry -notlike "") {
    "$LicenseKeyMExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyMExpiry"
}


#EMAIL--------------

$Email = if ($LicenseKeyNExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyNExpiry -and $LicenseKeyNExpiry -notlike "") {
    "$LicenseKeyNExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyNExpiry"
}


#Notes--------------

$Notes = if ($LicenseKeyOExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyOExpiry -and $LicenseKeyOExpiry -notlike "") {
    "$LicenseKeyOExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyOExpiry"
}


#Posts--------------

$Posts = if ($LicenseKeyPExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyPExpiry -and $LicenseKeyPExpiry -notlike "") {
    "$LicenseKeyPExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyPExpiry"
}


#Tasks--------------

$Tasks = if ($LicenseKeyTExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyTExpiry -and $LicenseKeyTExpiry -notlike "") {
    "$LicenseKeyTExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyTExpiry"
}


$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$answer = $wshell.Popup("Please Review your Add2Exchange License Expiration Dates. If your keys are expired please renew the software. Click OK to continue with the upgrade, or Cancel to Quit.

Expirations Dates as of $Today :

Calendars Sync= $Calendars
Contacts Sync=  $Contacts
Relationship Group Manager= $RGM
GAL Sync= $GAL
Mail Confidentiality= $Mail
Email Notifications= $Email
Notes Sync= $Notes
Posts Sync= $Posts
Tasks Sync= $Tasks


NOTE* Upgrading Add2Exchange with expired keys will stop synchronization!
", 0, "ATTENTION!! Add2Exchange Licensing", 0 + 1)
if ($answer -eq 2) { Break }




#Stop Menu Process
Stop-Process -Name "DidItBetterSupportMenu" -Force -ErrorAction SilentlyContinue

#Test for FTP

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
    Start-Process msiexec.exe -Wait -ArgumentList '/I "Add2Exchange_Upgrade.msi" /quiet' -ErrorAction Inquire -ErrorVariable InstallError;
    Write-Host "Finished...Upgrade Complete"

    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))



#Setting the Service Account Password
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password)
$SAP = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)

$SVC = Get-WmiObject win32_service -Filter "Name='Add2Exchange Service'"
$SVC.StopService();
$Result = $SVC.Change($Null, $Null, $Null, $Null, $Null, $Null, $Null, "$SAP")
If ($Result.ReturnValue -eq '0') { Write-Host "Add2Exchange Service Password Has Been Succsefully Updated" -ForegroundColor Green } Else { Write-Host "Error: $Result" }



#Setting the Add2Exchange Service to Delayed Start
Write-Host "Setting up Add2Exchange Service to Delayed Start"
sc.exe config "Add2Exchange Service" start= delayed-auto
Write-Host "Done"

#Starting Add2Exchange Service
Write-Host "Starting the Add2Exchange Service"
Start-Service -Name "Add2Exchange Service"
Write-Host "Done"

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting