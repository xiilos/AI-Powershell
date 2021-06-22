if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

$TestPath = Get-Itemproperty -path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name EnableADAL
$TestPath = Get-Itemproperty -path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name DisableADALatopWAMOverride
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "EnableADAL" -PropertyType DWORD -Value 0 -ErrorAction SilentlyContinue
  New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "DisableADALatopWAMOverride" -PropertyType DWORD -Value 1 -ErrorAction SilentlyContinue
}


Do {
  $confirmation = Read-Host "Would you like to Disable or Enable Modern Authentication [D/E]"
  if ($confirmation -eq 'D') {
    Write-Host "Disabling Modern Authentication"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "EnableADAL" -Type DWORD -Value 0 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "DisableADALatopWAMOverride" -Type DWORD -Value 1 -ErrorAction SilentlyContinue
    Write-Host "Done"
  }

    
  if ($confirmation -eq 'E') {
    Write-Host "Enabling Modern Authentication"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "EnableADAL" -Type DWORD -Value 1 -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name "DisableADALatopWAMOverride" -Type DWORD -Value 0 -ErrorAction SilentlyContinue
    Write-Host "Done"
  }
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting