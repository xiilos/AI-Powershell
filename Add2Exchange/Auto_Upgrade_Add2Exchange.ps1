<#
        .SYNOPSIS
        Automatically upgrades Add2Exchange to the newest version

        .DESCRIPTION
        Check and Creates scheduled update for Add2Exhcange
        Checks for outdated license keys and prompts before upgrading
        Downloads from S3
        Upgrades Add2Exchange to latest build
        Sets password for Add2Exchange Service
        Start Add2Exchange Console

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

#Checking Task Creation
Write-Host "Checking for Auto update Task..."
$TaskName = "Scheduled Update Add2Exchange"
$TaskExists = Get-ScheduledTask | Where-Object { $_.TaskName -like $TaskName }

If ($TaskExists) {
    Write-Host "Task Exists" 
} 

Else {
    $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
    Set-Location $Location
    #$Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1)
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Scheduled_Update_Add2Exchange.ps1"'
    $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
    Register-ScheduledTask -Action $Action -RunLevel Highest -TaskName "Scheduled Update Add2Exchange" -Description "Updates Add2Exchange to the latest Version" -User $UserID -Password $Password
    Write-Host "Done" 
}


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
$answer = $wshell.Popup("Please Review your Add2Exchange License Expiration Dates below.

If your Free Upgrade Period (Software Assurance) has passed and they are expired, you will need to renew the software to continue using it.
If you have already renewed or in that process, please Press Continue.

Your Software Mainenance Free Upgrade Periods

$Calendars for Calendars Synchronization
$Contacts for Contacts Synchronization
$RGM  for the Relationship Group Manager 
$GAL for GAL Synchronization 
$Notes for Notes Synchronization 
$Posts  for Posts Synchronization 
$Tasks for Tasks Synchronization 
$Mail for Email Synchronization
$Email  for Confidential Email Notifier

$Today is today’s date, Any keys for your active functioning modules will expire AFTER this date.  

Click OK to continue with your upgrade, or Cancel to Quit.

NOTE* Upgrading Add2Exchange with expired keys and outside Software Assurance will stop synchronization until a renewal is purchased and new keys are issued.

", 0, "ATTENTION!! Add2Exchange Licensing", 0 + 1)
if ($answer -eq 2) { Break }


#Stop Menu Process
Stop-Process -Name "DidItBetter Support Menu" -Force -ErrorAction SilentlyContinue


#Create zLibrary
Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\Add2Exchange Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Upgrades Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Add2Exchange Upgrades"
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

#Downloading Add2Exchange
Write-Host "Downloading Add2Exchange"
Write-Host "Please Wait......"


$bucketUrl = "https://s3.amazonaws.com/dl.diditbetter.com/"
$partialFileName = "a2e-enterprise_upgrade"

# web request to get the contents of the bucket
$request = [System.Net.WebRequest]::Create($bucketUrl)
$response = $request.GetResponse()

# Read the response stream
$reader = New-Object System.IO.StreamReader($response.GetResponseStream())
$contents = $reader.ReadToEnd()

# Extract the keys from the response
$keyPattern = "(?<=\<Key\>)[^<]+(?=\<\/Key\>)"
$keys = [regex]::Matches($contents, $keyPattern) | ForEach-Object { $_.Value }

# Find the first key that matches the partial filename
$matchingKey = $keys | Where-Object { $_ -like "$partialFileName*" } | Select-Object -First 1

# Download the matching file
$matchingFileUrl = "$bucketUrl$matchingKey"
$destinationPath = "C:\zlibrary\Add2Exchange Upgrades"
$fileName = "a2e-enterprise_upgrade.exe"
$downloadedfileName = [System.IO.Path]::GetFileName($matchingFileUrl)
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $matchingFileUrl -OutFile "$destinationPath\$fileName"

Write-Host "Finished Downloading"

#Upgrading from version_x.x to version_x.x
$CurrentVersionDB = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "CurrentVersionDB" -ErrorAction SilentlyContinue
Write-Host "You are now upgrading Add2Exchange Enterprise from Version $CurrentVersionDB to $downloadedfileName" -ForegroundColor Green

#Stop Add2Exchange Service
Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 5
Write-Host "Done"

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


#Remove Add2Exchange
Write-Host "Removing Add2Exchange"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
$Program.Uninstall()
Write-Host "Done"


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

#Starting Add2Exchange Console
Write-Host "Starting the Add2Exchange Console"
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
Push-Location $Install
Start-Process "./Console/Add2Exchange Console.exe"

#Start-Service -Name "Add2Exchange Service"
Write-Host "Done"



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting
