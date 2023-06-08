<#
        .SYNOPSIS
        

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


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

# Set the S3 bucket URL and file path
$bucketUrl = "https://s3.amazonaws.com/<your-bucket-name>"
$filePath = "C:\path\to\your\file.txt"

# Read the file content
$fileContent = Get-Content -Path $filePath -Raw

# Generate the S3 object key based on the file name
$fileName = [System.IO.Path]::GetFileName($filePath)
$s3ObjectKey = "files/$fileName"

# Set the S3 object URL
$s3ObjectUrl = $bucketUrl + "/" + $s3ObjectKey

# Set the HTTP headers
$headers = @{
    "Content-Type" = "application/octet-stream"
}

# Make the HTTP PUT request to upload the file
Invoke-RestMethod -Uri $s3ObjectUrl -Method Put -Headers $headers -Body $fileContent

# Output the S3 object URL
Write-Host "File uploaded successfully."
Write-Host "S3 Object URL: $s3ObjectUrl"


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting