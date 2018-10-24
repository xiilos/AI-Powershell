if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}
# Execution Policy

Write-Host "Chekcing version of powershell"
$PSVersionTable.PSVersion

$confirmation = Read-Host "Would you like me to update powershell? [Y/N]"
if ($confirmation -eq 'y') {
  

Write-Host "Updating Powershell"
Write-Host "Downloading the latest powershell Update"
$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/W2K12-KB3191565-x64.msu"
$output = "$PSScriptRoot\Powershell_5.msu"
$start_time = Get-Date

$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $output)
#OR
(New-Object System.Net.WebClient).DownloadFile($url, $output)

Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

}


$confirmation = Read-Host "Would you like me to Install MSonline Module? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Azure MSonline module"
Install-Module MSonline -Confirm:$false

}



Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting