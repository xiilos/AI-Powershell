<#
        .SYNOPSIS
        Automatically upgrades Add2Outlook to the newest version

        .DESCRIPTION
        Check and Creates scheduled update for Add2Outlook
        Checks for outdated license keys and prompts before upgrading
        Downloads from S3
        Upgrades Add2Outlook to latest build
        Sets password for Add2Outlook Service
        Start Add2Outlook interface

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

#Stop Menu Process
Stop-Process -Name "DidItBetterSupportMenu" -Force -ErrorAction SilentlyContinue

#Create zLibrary\Add2Outlook Toolkit Directory
Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\A2OToolKit"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Outlook Toolkit Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\A2OToolKit"
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


#Downloading Add2Outlook Toolkit
Write-Host "Downloading Add2Outlook Toolkit"
Write-Host "Please Wait......"

# Replace the value of $bucketUrl with the public Amazon S3 URL of your bucket
$bucketUrl = "https://s3.amazonaws.com/dl.diditbetter.com/"

# Replace the value of $partialFileName with the first part of the filename you know
$partialFileName = "Add2Outlook Toolkit Full Installation"

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
$destinationPath = "C:\zlibrary\A2OToolKit"
$fileName = "Add2OutlookToolkitFullInstallation.exe"
#$downloadedfileName = [System.IO.Path]::GetFileName($matchingFileUrl)
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $matchingFileUrl -OutFile "$destinationPath\$fileName"

Write-Host "Finished Downloading"


#Unpacking Add2Outlook Toolkit
Write-Host "Unpacking Add2Outlook Toolkit"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\A2OToolKit"
Write-Host "Done"

#Installing Add2Outlook Toolkit
Write-Host "Installing Add2Outlook Toolkit"
Start-Process -FilePath "./Add2OutlookToolkitFullInstallation.exe" -wait -ErrorAction Stop
Write-Host "Finished...Upgrade Complete"
Start-Sleep -Seconds 2

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting