if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Adding the WU Module in Powershell
Import-Module "PSWindowsUpdate" -ErrorAction SilentlyContinue
If ($error) {
    Write-Host "Adding Windows Update Module"
    Install-Module PSWindowsUpdate -Confirm:$False
} 
Else {
    Write-Host 'Module is installed'
    Update-Module "PSWindowsUpdate"
}


#Installing Windows Updates
Write-Host "Getting Windows Updates...Please Wait..."
Get-WindowsUpdate -Install -Confirm:$False –IgnoreReboot


#Cleaning
Write-Host "Cleaning Up....."
Write-Host "Flushing DNS"
IPConfig /FlushDNS
Write-Host "Done"
Write-Host "Cleaning up Files via Disk Cleanup"
Cleanmgr /verylowdisk /SageRun:5 | Out-Null
Write-Host "Done"

#Rebooting
Write-Host "Rebooting...."
Shutdown -r -t 10



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting