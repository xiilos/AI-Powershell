if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

# Bypass AutoDiscover

$Val = Get-ItemProperty -Path "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -ErrorAction SilentlyContinue -ErrorVariable Error

If ($Error) {
    New-ItemProperty "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -Value 1 -PropertyType "DWord"
}

if($val.ExcludeExplicitO365Endpoint -ne 0)

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


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting