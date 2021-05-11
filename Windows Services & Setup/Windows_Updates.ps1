if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue


#Variables
$Nodes = $env:COMPUTERNAME


#Task Creator

if (Get-ScheduledTask "PS Windows Updates" -ErrorAction SilentlyContinue) {
    Write-Host "PS Windows Updates Task Already Exists..."
}

else {
    $Location = "\\fileserv-db1\Logs\PowerShell Files\"
    $Trigger= New-JobTrigger -Weekly -DaysOfWeek Saturday -At 12am
    $Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Windows_Updates.ps1"'
    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "PS Windows Updates" -Description "Powershell Windows Updates once a Week"
    Write-Host "Task Created"
}




# Script #
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

#Adding the WU Module in Powershell
Install-PackageProvider -Name NuGet -Force
Import-Module "PSWindowsUpdate" -ErrorAction SilentlyContinue
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue
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
Get-WindowsUpdate -MicrosoftUpdate -Install -AcceptAll -AutoReboot -Confirm:$false | Out-File "\\fileserv-db1\Logs\MS Update Logs\$Nodes-$(Get-Date -f yyyy-MM-dd)-MSUpdates.log" -Force



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting