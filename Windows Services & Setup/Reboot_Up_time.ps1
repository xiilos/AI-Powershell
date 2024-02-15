<#
        .SYNOPSIS
        

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         Genesis

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#OpenVPN
Start-Process -FilePath "C:\Program Files\OpenVPN\bin\openvpn-gui.exe" -ArgumentList "--connect HIV-Main-UDP4-1194-www.diditbetter.com-config.ovpn"
Start-Sleep -s 5

#Open Outlook
Start-Process outlook
Start-Sleep -s 5

#Visual Studio
start-process code -verb RunAs 
Start-Sleep -s 2

#mremote
Start-Process "\\FILE_SERVER\Sandbox\Port Apps\PortableApps\mRemoteNG-Portable-1.76\mremoteng.exe"
Start-Sleep -s 2

#Morning Startup
Start-Process -FilePath "$env:USERPROFILE\Desktop\Morning Startup.lnk"
Start-Sleep -s 2


#Daily work Files
$files = @(
    "\\file_server\Work\Advantage International\Morning startup\daily notes.txt",
    "\\file_server\Work\Advantage International\Morning startup\powershell to do.txt",
    "\\file_server\Work\Advantage International\Morning startup\infrastructure to do.txt"
)

foreach ($file in $files) {
    Start-Process -FilePath "notepad.exe" -ArgumentList $file
}
Start-Sleep -s 2

#Weekly Time Sheet
Start-Process "\\file_server\Work\Advantage International\Morning startup\Weekely Time Sheet\Weekly Time Sheet.xls"

#Keepass
Start-Process "\\FILE_SERVER\Sandbox\Port Apps\PortableApps\KeePass\keepass.exe"
Start-Sleep -s 2


# PowerShell script to open multiple websites in new tabs
# Define Personal URLs to open
$urls = @(
    "https://feedly.com/",
    "https://protonmail.com/"
)

# Loop through each URL and open it in the default browser
foreach ($url in $urls) {
    Start-Process "brave.exe" $url
}

# Define Work URLs to open
$urls = @(
    "https://join.diditbetter.com",
    "http://issues.diditbetter.com",
    "https://viz.greynoise.io/",
    "https://chat.openai.com/",
    "https://www.diditbetter.com/dibai/"
)

# Loop through each URL and open it in the default browser
foreach ($url in $urls) {
    Start-Process "brave.exe" $url
}
pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting