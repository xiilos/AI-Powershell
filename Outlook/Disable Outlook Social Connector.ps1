if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Connect to Outlook
$outlook = New-Object -ComObject Outlook.Application

# Get the Outlook Social Connector add-in
$oscAddin = $outlook.COMAddIns | Where-Object {$_.ProgId -eq "Outlook.SocialConnector.Addin"}

if ($oscAddin -ne $null)
{
    # Disable the add-in
    $oscAddin.Connect = $false
    $oscAddin.Update()
    
    Write-Host "Outlook Social Connector add-in has been disabled."
}
else
{
    Write-Host "Outlook Social Connector add-in is not installed."
}

# Release the Outlook COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting