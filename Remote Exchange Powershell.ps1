$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session