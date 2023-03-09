<#
        .SYNOPSIS
        Automatically upgrades Recovery and Migration Manager to the newest version

        .DESCRIPTION
        Check and Creates scheduled update for RMM
        Checks for outdated license keys and prompts before upgrading
        Downloads from S3
        Upgrades RMM to latest build
        Start RMM interface

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

$Today  is today’s date, and any keys for your active functioning modules will expire AFTER this date.

Click OK to continue with your upgrade, or Cancel to Quit.

NOTE* Upgrading Recovery and Migration Manager with expired keys and outside Software Assurance will stop synchronization until a renewal is purchased and new keys are issued.

", 0, "ATTENTION!! RMM Licensing", 0 + 1)
if ($answer -eq 2) { Break }



#Create zLibrary\RMM Directory
Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\RMM Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "RMM Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\RMM Upgrades"
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

#Downloading RMM
Write-Host "Downloading Recovery and Migration Manager"
Write-Host "Please Wait......"

# Replace the value of $bucketUrl with the public Amazon S3 URL of your bucket
$bucketUrl = "https://s3.amazonaws.com/dl.diditbetter.com/"

# Replace the value of $partialFileName with the first part of the filename you know
$partialFileName = "rmm-enterprise"

# Create a web request to get the contents of the bucket
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
$destinationPath = "C:\zlibrary\RMM Upgrades"
$fileName = "rmm-enterprise.exe"
#$downloadedfileName = [System.IO.Path]::GetFileName($matchingFileUrl)
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $matchingFileUrl -OutFile "$destinationPath\$fileName"

Write-Host "Finished Downloading"


#Remove Recovery and Migration Manager
Write-Host "Removing Recovery and Migration Manager"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Recovery and Migration Manager'"
$Program.Uninstall()
Write-Host "Done"

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