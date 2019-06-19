if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
Get-Service | Where-Object { $_.DisplayName –like “Microsoft Exchange *” } | Set-Service –StartupType Automatic

Get-Service | Where-Object { $_.DisplayName –eq “IIS Admin Service” } | Set-Service –StartupType Automatic

Get-Service | Where-Object { $_.DisplayName –eq “IIS Admin Service” } | Start-Service

Get-Service | Where-Object { $_.DisplayName –like “Microsoft Exchange *” } | Start-Service


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting