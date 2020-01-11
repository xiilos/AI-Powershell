if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

$dir = “c:\temp”

mkdir $dir

$webClient = New-Object System.Net.WebClient

$url = “https://go.microsoft.com/fwlink/?LinkID=799445"

$file = “$($dir)\Win10Upgrade.exe”

$webClient.DownloadFile($url,$file)

Start-Process -FilePath $file -ArgumentList “/quietinstall /skipeula /auto upgrade /copylogs $dir” 

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting