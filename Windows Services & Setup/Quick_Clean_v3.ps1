if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue


# Script #
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

#Adding the WU Module in Powershell
Install-PackageProvider -Name NuGet -Force
Import-Module "PSWindowsUpdate" -ErrorAction SilentlyContinue
If ($error) {
    Write-Host "Adding Windows Update Module"
    Install-Module PSWindowsUpdate -Confirm:$False -Force
} 
Else {
    Write-Host 'Module is installed... Updating Current Module'
    Update-Module "PSWindowsUpdate"
    Write-Host "Done"
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
cleanmgr /sageset:1
Pause
cleanmgr /sagerun:1
Write-Host "Running Cleaner. Hit Enter When Finished"

Pause

#Rebooting
Write-Host "Rebooting...."
Shutdown -r -t 10



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting