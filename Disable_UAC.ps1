if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


# Disable UAC

Write-Host "Disabling UAC In the Registry"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null

Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Please Reboot",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting