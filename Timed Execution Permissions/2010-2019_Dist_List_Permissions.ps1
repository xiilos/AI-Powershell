if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Name.txt"
$ServiceAccount = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
$Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring
$DistributionGroupName = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dist_List_Name.txt"

$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to Distribution Lists
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName) {
    Add-MailboxPermission -Identity $Member.name -User $ServiceAccount -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}


Get-PSSession | Remove-PSSession
Exit

# End Scripting