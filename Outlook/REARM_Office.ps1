if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Detect Bitness
$64Bits = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook" -Name "Bitness" | Select-Object Bitness -ExpandProperty Bitness -ErrorAction SilentlyContinue

If ($64Bits -eq 'x64'){
  Set-Location "C:\Program Files\Microsoft Office\Office16"

.\OSPPREARM.EXE

cscript .\ospp.vbs /dstatus

Pause
}

$32Bits = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook" -Name "Bitness" | Select-Object Bitness -ExpandProperty Bitness -ErrorAction SilentlyContinue

If ($32Bits -eq 'x86'){
  Set-Location "C:\Program Files (x86)\Microsoft Office\Office16"

.\OSPPREARM.EXE

cscript .\ospp.vbs /dstatus

Pause
}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting