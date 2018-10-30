if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Unrestricted

# Remove Outlook Add-ins

Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALCONNECTOR.DLL*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALCONNECTOR.DLL','SOCIALCONNECTORbckup.dll' }
Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALPROVIDER.DLL*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALPROVIDER.DLL','SOCIALPROVIDERbckup.dll' }
Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS" "*ColleagueImport.dll*" -Recurse | Rename-Item -NewName {$_.name -replace 'ColleagueImport.dll','ColleagueImportbckup.dll' }

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting