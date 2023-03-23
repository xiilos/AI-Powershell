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


# Get current version of software
$version = "1.0.0" # Replace with your actual version number
$versionParts = $version.Split('.')

# Generate serial key
$serialKey = ""
for ($i = 0; $i -lt $versionParts.Count; $i++) {
    # Convert version part to integer
    $partValue = [int]$versionParts[$i]

    # Add value to serial key
    $serialKey += [char]($partValue + 65)
}

# Output serial key
Write-Output "Serial key for version $version : $serialKey"









# Define the version number
$version = "1.0.0"

# Generate a random 4-character string
$randomString = -join ((65..90) + (97..122) | Get-Random -Count 4 | ForEach-Object {[char]$_})

# Combine the version number and random string to create the serial key
$serialKey = $version + "-" + $randomString

# Output the serial key
Write-Output $serialKey


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting