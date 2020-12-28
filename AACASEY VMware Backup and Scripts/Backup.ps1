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
Start-Sleep -S 10

Write-Host "Shutting Down SCO-UNIX"

Invoke-sshcommand -index 0 -command "shutdown -y -i0 -g0"


Start-Process Powershell .\Stop-SCO-UNIX.bat
Remove-SSHSession -Index 0 -Verbose
Start-Sleep -S 10


#Write-Host "Shutting Down VMware Workstation"
Write-Host "Waiting 550 seconds for Virtual Machines to Power Off.........."
Start-Sleep -S 550
#Stop-Process -Name vmware



#Back Up Files
Write-Host "Backing up SCO-Unix. Please wait....."
Copy-Item 'E:\vmware\SCO-Unix' -Destination ('\\SEAGATE-D2\VMware\VMware Backups\SCO-Unix\' + (get-date -Format MM-dd-yyyy)) -Recurse

Write-Host "Backing up SBS 2011. Please wait....."
Copy-Item 'E:\vmware\SBS2' -Destination ('\\SEAGATE-D2\VMware\VMware Backups\SBS\' + (get-date -Format MM-dd-yyyy)) -Recurse

Write-Host "Finished Backing up All Files"
Start-Sleep -S 10

#Compress-Archive 'C:\Virtual Machines\SCO-Unix' -DestinationPath ('C:\Backup\SCO-Unix\' + (get-date -Format MM-dd-yyyy) + '.zip')



#Starting Virtual Machines

Write-Host "Starting Up Virtual Machines"

Start-Process Powershell .\Start-SCO-UNIX.bat
Start-Sleep -S 10
Start-Process Powershell .\Start-SBS.bat


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting