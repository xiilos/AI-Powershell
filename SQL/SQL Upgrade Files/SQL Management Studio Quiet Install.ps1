<#
        .SYNOPSIS
        Downloads and Silently Installs SQL Managament Studio 19x

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #


#Download SQL Management Studio
Write-Host "Downloading SQL Management Studio"
$URL = "https://s3.amazonaws.com/dl.diditbetter.com/SQL%20Express/SSMS-Setup-ENU.exe"
$Output = "C:\zlibrary\SQL Upgrade\SSMS-Setup-ENU.exe"
$Start_Time = Get-Date
    
    (New-Object System.Net.WebClient).DownloadFile($URL, $Output)
    
Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"
    
Write-Host "Finished Downloading"

# Define the path to the SSMS setup executable
$ssmsSetupPath = "C:\zlibrary\SQL Upgrade\SSMS-Setup-ENU.exe"

# Start the installation process quietly
Write-Host "Installing SQL Management Studio. Please Wait..."
Start-Process -FilePath $ssmsSetupPath -ArgumentList "/install /quiet /norestart" -Wait

# Check for successful installation (optional)
if ($LASTEXITCODE -eq 0) {
    Write-Host "SSMS was successfully installed."
}
else {
    Write-Host "There was an error during the SSMS installation."
    
}

Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting