if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}



# Script #



#Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3

#Create zLibrary

$TestPath = "C:\zlibrary"

if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "zLibrary exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path C:\zlibrary
 }

#Download the latest Full Installation
$confirmation = Read-Host "Would you like me to download the latest Add2Exchange? [Y/N]"
    if ($confirmation -eq 'y') {

        Write-Host "Downloading Add2Exchange......"
        Write-Host "This can take a few Minutes"
        $url = "ftp://ftp.diditbetter.com/A2E-Enterprise/A2ENewInstall/a2e-enterprise.zip"
        $output = "C:\zlibrary\a2e-enterprise.zip"
        (New-Object System.Net.WebClient).DownloadFile($url, $output)
        Write-Host "Download Done"
        Write-Host "Unpacking"    
        Expand-Archive -Path "C:\zlibrary\a2e-enterprise.zip" -DestinationPath "c:\zlibrary"

}

# End Scripting