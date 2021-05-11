if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue


#Variables


# Script #

Remove-Item -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Reliability\* -Recurse #remove all keys within this folder




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting