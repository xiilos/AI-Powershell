<#
        .SYNOPSIS
        

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


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

$pass = Read-Host -AsSecureString "Enter password starting with !"
$plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))

if ($plainPass.StartsWith("!")) {
    Write-Host "Password accepted"
} else {
    Write-Host "Password must start with !"
}



#also -raw at end gets raw info





$password = Read-Host "Enter a password that starts with !:" -AsSecureString
$passwordString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)) #this makes it readable
if ($passwordString.StartsWith("!")) {
    Write-Host "Password accepted."
} else {
    Write-Host "Password must start with !"
}











Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting