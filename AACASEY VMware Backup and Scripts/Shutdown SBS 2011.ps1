if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Install and Check Posh-SSH Module Status

$error.clear()
        Import-Module "Posh-SSH" -ErrorAction SilentlyContinue
        If ($error) {
            
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PSRepository -Name psgallery -InstallationPolicy Trusted
            Install-Module Posh-SSH -Confirm:$false -WarningAction "Inquire"
        } 


# Logging into SCO-UNIX

New-SshSession -Computername "10.200.200.10" -Credential root -AcceptKey -ErrorAction SilentlyContinue -Errorvariable LoginError

If ($LoginError) {
            
  New-SshSession -Computername "10.200.200.10" -Credential root -AcceptKey
} 


#Stopping Virtual Machines

Write-Host "Shutting Down SBS 2011"
Start-Process Powershell .\Stop-SBS.bat
Start-Sleep -S 60

#Write-Host "Shutting Down SCO-UNIX"
#Invoke-sshcommand -index 0 -command "shutdown -y -i0 -g0"


#Start-Process Powershell .\Stop-SCO-UNIX.bat
#Remove-SSHSession -Index 0 -Verbose
#Start-Sleep -S 60


Write-Host "Shutting Down VMware Workstation"
Write-Host "Waiting 900 seconds for Virtual Machines to Power Off.........."
Start-Sleep -S 900
Stop-Process -Name vmware


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting