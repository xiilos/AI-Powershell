if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Unrestricted


Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites]
"Session Manager"="Computer\\HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager"
"enablelua"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System"
".NETFramework"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\.NETFramework"
"OpenDoor Software®"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\OpenDoor Software®"
"Add2Exchange"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\OpenDoor Software®\\Add2Exchange"
"pendingfilerename"="Computer\\HKEY_LOCAL_MACHINE\\SYSTEM\\ControlSet001\\Control\\Session Manager"
"AutoDiscover"="Computer\\HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office"
"Office"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office"
"Group Policy History"="Computer\\HKEY_CURRENT_USER\\Software\Microsoft\\Windows\\CurrentVersion\\Group Policy\\History"
"Windows Logon"="Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon"



Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting