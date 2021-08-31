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

$User = "root"
$Password = "fudge"
$Computer = "10.200.200.10"

$secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential($User, $secpasswd)

New-SshSession -Computername $Computer -Credential $Credentials -AcceptKey -ErrorAction SilentlyContinue -Errorvariable LoginError
  
If ($LoginError) {
              
    New-SshSession -Computername $Computer -Credential $Credentials -AcceptKey
} 
  
#Stopping Virtual Machines
  
Write-Host "Shutting Down SCO-UNIX"
  
Invoke-sshcommand -index 0 -command "shutdown -y -i0 -g0"
  
  
Start-Process Powershell .\Stop-SCO-UNIX.bat
Remove-SSHSession -Index 0 -Verbose
Start-Sleep -S 10
  
  
#Write-Host "Shutting Down VMware Workstation"
Write-Host "Waiting 90 seconds for Virtual Machines to Power Off.........."
Start-Sleep -S 90

  
#Backup Retention (3)
$Path = "S:\SCO-Unix"
Push-Location $Path

$FolderCount = (Get-ChildItem -Path $Path | where-object { $_.PSIsContainer }).Count
If ($FolderCount -gt 3) {
    $Object = Get-Childitem -Path $Path | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -Last 1
    Write-Host "Removing $Object"
    Remove-Item $Object -Recurse -Force
    Write-Host "Done"
}

  
#Back Up Files
Write-Host "Backing up SCO-Unix. Please wait....."
Copy-Item 'D:\VMWare_Workstation\SCO-Unix' -Destination ('S:\SCO-Unix\' + (get-date -Format MM-dd-yyyy)) -Recurse
Write-Host "Finished Backing up All Files"
Start-Sleep -S 10

#Starting Virtual Machines
  
Write-Host "Starting Up Virtual Machines"
Push-Location "D:\zLibrary\VMware Backup Scripts"  
Start-Process Powershell .\Start-SCO-UNIX.bat
Start-Sleep -S 10
  
  
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
  
# End Scripting