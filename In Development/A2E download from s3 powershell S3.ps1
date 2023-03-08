if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

# Replace the value of $bucketUrl with the public Amazon S3 URL of your bucket
$bucketUrl = "https://s3.amazonaws.com/downloads.diditbetter.com/"

# Replace the value of $partialFileName with the first part of the filename you know
$partialFileName = "a2e-enterprise_upgrade"

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
$destinationPath = "C:\zlibrary\Add2Exchange Upgrades"
$fileName = [System.IO.Path]::GetFileName($matchingFileUrl)
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $matchingFileUrl -OutFile "$destinationPath\$fileName"



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting