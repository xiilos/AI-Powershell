if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

# Script #

$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$answer = $wshell.Popup("Before you run this script, you must follow the steps from the below link before proceeding.
https://docs.microsoft.com/en-us/powershell/exchange/mfa-connect-to-exchange-online-powershell?view=exchange-ps
", 0, "ATTENTION!! Office 365 MFA Login", 0 + 1)
if ($answer -eq 2) { Break }

Write-Host "Connecting to Office 365 with MFA"

$Username = Read-Host "Please type in the username that you will use to connect to Office 365 with MFA. Example: Tom@contoso.com"

    #Office 365--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Connect-EXOPSSession -UserPrincipalName $Username

    

# End Scripting