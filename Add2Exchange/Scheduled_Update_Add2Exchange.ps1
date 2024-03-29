<#
        .SYNOPSIS
        Scheduled Task to automatically upgrade Add2Exchange to the newest version

        .DESCRIPTION
        Downloads from S3
        Upgrades Add2Exchange to latest build
        Sets password for Add2Exchange Service
        Starts Add2Exchange Service after succesfull install

        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

        To run this as a scheduled task open CMD prompt and type in: schtasks /run /tn "Scheduled Update Add2Exchange" 
        Note* the task must be already created before running this in CMD
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
    Write-Host "Connection failed! Quitting"
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
$CurrentVersionDB = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" -Name "CurrentVersionDB" -ErrorAction SilentlyContinue
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
Write-Host "Installing Add2Exchange"
$Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Push-Location $Location
Start-Process msiexec.exe -Wait -ArgumentList '/I "Add2Exchange_Upgrade.msi" /quiet' -ErrorAction Inquire -ErrorVariable InstallError;
Write-Host "Finished...Upgrade Complete"

If ($InstallError) { 
    Write-Warning -Message "Something Went Wrong with the Install!"
    Write-Host "Try Installing Manually"
    Start-Sleep -S 2
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}




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

#Starting Add2Exchange
Write-Host "Starting the Add2Exchange Service"
Start-Sleep -Seconds 3
Start-Service -Name "Add2Exchange Service"
Write-Host "Done"



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting

