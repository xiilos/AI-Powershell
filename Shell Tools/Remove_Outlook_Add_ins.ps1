if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
  
#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass
  
#Variables

# Remove Outlook Social Connector
  
$TestPath = "C:\Program Files (x86)\Microsoft Office\root\Office16"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    Write-Host "32bit Outlook"
    Write-Host "Setting Load Behavior to 0"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

}
Else {
    Write-Host "64bit Outlook"
    Write-Host "Setting Load Behavior to 0"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

}
  
Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Restart Outlook", 0, "Done", 0x1)
Start-Sleep 1
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
  
# End Scripting