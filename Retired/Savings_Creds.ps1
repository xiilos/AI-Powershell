if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}



# Script #

$error.clear()
Import-Module -Name "CredentialManager" -ErrorAction SilentlyContinue
If($error){Write-Host "Adding Cred Manager Module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module "CredentialManager" -Confirm:$false -WarningAction "Inquire"} 
Else{Write-Host 'Credential Manager Module is already installed'}

$Target = Read-host "Exchange Servername?"
$UserName = Read-host "Admin Username?"
$Secure = Read-host "Password?" -AsSecureString
New-StoredCredential -Target $Target -UserName $UserName -SecurePassword $Secure -Persist LocalMachine -Type Generic



Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting