if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


# Disable UAC

$Val = Get-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA"

if($val.EnableLUA -ne 0)

{
Set-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA" -value 0
Write-Host "UAC is now Disabled"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Reboot",0,"Done",0x1)
}

Else {

  $wshell = New-Object -ComObject Wscript.Shell
  $wshell.Popup("UAC Is already Disabled",0,"Done",0x1)

}

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting