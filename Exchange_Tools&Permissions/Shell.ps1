<#
        .SYNOPSIS
        Shell

        .DESCRIPTION
        Simple open another PS session to shell into Exchange of Office365
        Calls another powershell file "Shell into Exchange"


        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
Powershell.exe -noexit ".\Shell_Into_Exchange.ps1" -noprofile

# End Scripting
