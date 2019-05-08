if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
  
#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass
  
#Variables
$DateStamp = =Get-Date -uformat "%Y-%m-%d@%H-%M-%S" 
# Remove Outlook Add-ins
  
$TestPath = "C:\Program Files (x86)\Microsoft Office\root\Office16"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    Write-Host "32bit Outlook"
    Remove-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALCONNECTORbckup.dll" -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALPROVIDERbckup.dll" -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\ColleagueImportbckup.dll" -ErrorAction SilentlyContinue
  
  
    Rename-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALCONNECTOR.dll" -NewName "SOCIALCONNECTOR.Backup$DateStamp" -ErrorAction SilentlyContinue
    Rename-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALPROVIDER.dll" -NewName "SOCIALPROVIDER.Backup$DateStamp" -ErrorAction SilentlyContinue
    Rename-Item -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\ColleagueImport.dll" -NewName "ColleagueImport.Backup$DateStamp" -ErrorAction SilentlyContinue
}
Else {
    Write-Host "64bit Outlook"
    Remove-Item -Path "C:\Program Files\Microsoft Office\root\Office16\SOCIALCONNECTORbckup.dll" -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Program Files\Microsoft Office\root\Office16\SOCIALPROVIDERbckup.dll" -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Program Files\Microsoft Office\root\Office16\ADDINS\ColleagueImportbckup.dll" -ErrorAction SilentlyContinue
  
  
    Rename-Item -Path "C:\Program Files\Microsoft Office\root\Office16\SOCIALCONNECTOR.dll" -NewName "SOCIALCONNECTOR.Backup$DateStamp" -ErrorAction SilentlyContinue
    Rename-Item -Path "C:\Program Files\Microsoft Office\root\Office16\SOCIALPROVIDER.dll" -NewName "SOCIALPROVIDER.Backup$DateStamp" -ErrorAction SilentlyContinue
    Rename-Item -Path "C:\Program Files\Microsoft Office\root\Office16\ADDINS\ColleagueImport.dll" -NewName "ColleagueImport.Backup$DateStamp" -ErrorAction SilentlyContinue
}
  
Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Restart Outlook", 0, "Done", 0x1)
Start-Sleep 1
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
  
# End Scripting