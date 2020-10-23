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
Get-ItemProperty -Path "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -ErrorAction SilentlyContinue -ErrorVariable Error

    If ($Error) {
        New-Item "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook" -Name "Autodiscover"
    }


Do {
  $confirmation = Read-Host "Do you want to Bypass AutoDiscover and connect directly to Office 365 [Y/N]"
  if ($confirmation -eq 'Y') {
    Write-Host "Bypassing AutoDiscover for Office 365"
   
        New-ItemProperty "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -Value 1 -PropertyType "DWord"
    
    
    if($val.ExcludeExplicitO365Endpoint -ne 0)
    
    {
    Set-ItemProperty -Path "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -value 1
    }
    Write-Host "AutoDiscover has now been bypassed"
    Write-Host "Done"
  }

    
  if ($confirmation -eq 'N') {
    Write-Host "Disabling Bypass for Office 365"
    Set-ItemProperty -Path "HKCU:Software\Policies\Microsoft\Office\16.0\Outlook\Autodiscover" -Name "ExcludeExplicitO365Endpoint" -value 0
    Write-Host "AutoDiscover has now been set to default on premise"
    Write-Host "Done"
  }
  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting