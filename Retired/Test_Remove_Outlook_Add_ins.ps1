if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
reg.exe add "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\OscAddin.Connect" /t REG_DWORD /v LoadBehavior /d 0 /f


Write-Host "ttyl"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting