if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Variables #
$TenentUsername = Get-StoredCredential -target 'Office365_Tenent' -ascredentialobject | Select-Object Username -ExpandProperty Username
$TenentPassword = Get-StoredCredential -target 'Office365_Tenent' | Select-Object password -ExpandProperty password
$SyncUsername = Get-StoredCredential -target 'Sync_Account' | Select-Object Username -ExpandProperty Username
$DistributionGroupName = Get-StoredCredential -target 'Distribution_Group_Name' -ascredentialobject | Select-Object username -ExpandProperty username

# Script #

$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $TenentUsername, $TenentPassword

Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

#Timed Execution Permissions to Distribution Lists
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName) {
    Add-MailboxPermission -Identity $Member.name -User $SyncUsername -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}


Get-PSSession | Remove-PSSession
Exit

# End Scripting