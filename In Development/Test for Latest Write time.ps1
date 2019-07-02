if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
# Find latest installer
$url = 'http://dl.diditbetter.com/'
$site = Invoke-WebRequest -UseBasicParsing -Uri $url
$table = $site.links | Where-Object{ $_.tagName -eq 'a2e-enterprise' -and $_.href.ToLower().Contains('upgrade') -and $_.href.ToLower().EndsWith("exe") } | Sort-Object href -desc | Select-Object href -first 1
$filename = $table

# Download installer
$src = $url + $filename
$dst = "c:\downloads" + $filename
Invoke-WebRequest $src -OutFile $dst


Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting