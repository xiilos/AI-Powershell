if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Re-Enable Outlook Add-ins

$TestPath = "C:\Program Files (x86)\Microsoft Office\root\Office16"

if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
Write-Host "32bit Outlook"
Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALCONNECTOR.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALCONNECTOR.Backup','SOCIALCONNECTOR.dll' }
Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALPROVIDER.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALPROVIDER.Backup','SOCIALPROVIDER.dll' }
Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS" "*ColleagueImport.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'ColleagueImport.Backup','ColleagueImport.dll' }
 }
Else {
Write-Host "64bit Outlook"
Get-ChildItem -path "C:\Program Files\Microsoft Office\root\Office16" "*SOCIALCONNECTOR.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALCONNECTOR.Backup','SOCIALCONNECTOR.dll' }
Get-ChildItem -path "C:\Program Files\Microsoft Office\root\Office16" "*SOCIALPROVIDER.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALPROVIDER.Backup','SOCIALPROVIDER.dll' }
Get-ChildItem -path "C:\Program Files\Microsoft Office\root\Office16\ADDINS" "*ColleagueImport.Backup*" -Recurse | Rename-Item -NewName {$_.name -replace 'ColleagueImport.Backup','ColleagueImport.dll' }
 }

Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Restart Outlook",0,"Done",0x1)
Start-Sleep 7
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting