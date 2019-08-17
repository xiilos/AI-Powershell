if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Variables #
$ExchangeUsername = Get-StoredCredential -target 'Exchange_Server' -ascredentialobject | Select-Object Username -ExpandProperty Username
$ExchangePassword = Get-StoredCredential -target 'Exchange_Server' | Select-Object password -ExpandProperty password
$SyncUsername = Get-StoredCredential -target 'Sync_Account' | Select-Object Username -ExpandProperty Username
$Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchangename.txt"
$DistributionGroupName = Get-StoredCredential -target 'Distribution_Group_Name' -ascredentialobject | Select-Object username -ExpandProperty username

# Script #

$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Exchangeusername, $ExchangePassword

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to Distribution Lists
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName) {
    Add-MailboxPermission -Identity $Member.name -User $SyncUsername -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}


Get-PSSession | Remove-PSSession
Exit

# End Scripting