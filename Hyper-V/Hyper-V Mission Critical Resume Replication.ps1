if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Notes: Resumes and Resyncs all Missions Critical VMs

# Script #

Resume-VMReplication –VMName "3cx-phone", "DIBDC", "DIBDC1", "DIBEX10", "Docker-a", "IIS", "MOJO-B", "Redmine", "SpamTitan", "Spree-B", "SQL", "Veeam" -Resynchronize

Write-Host "Done"
Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
