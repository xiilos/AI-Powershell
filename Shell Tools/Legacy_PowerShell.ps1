<#
        .SYNOPSIS
        Powershell script to check and update .net and 

        .DESCRIPTION
        Will update Powershell to 5.1 if .net is below version 4.5
        Get Windows Version
        Get .net version


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

# Check if .Net 4.5 or above is installed
$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389)) {
    Write-Host "We need to download .Net 4.5.2"
    Write-Host "Downloading"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    $output = "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)    
    Start-Process -FilePath "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing and run this again", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}


#Check Operating Sysetm
$BuildVersion = [System.Environment]::OSVersion.Version


#OS is 10+
if ($BuildVersion.Major -like '10') {
    Write-Host "WMF 5.1 is not supported for Windows 10 and above"
        
}

#OS is 7
if ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1') {
        
    Write-Host "Downloading WMF 5.1 for 7+"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
    $output = "C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Start-Process -FilePath 'C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu' -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#OS is 8
elseif ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3') {
        
    Write-Host "Downloading WMF 5.1 for 8+"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
    $output = "C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Start-Process -FilePath 'C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu' -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}


Write-Host "Nothing to do"
Write-Host "You Are on the latest version of PowerShell"
Write-Host "Quitting"
Start-Sleep 2
Get-PSSession | Remove-PSSession
Exit

# End Scripting