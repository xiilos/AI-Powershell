if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Variables #

$User = "Type the Name of Sync Service account here"

#NOTE* For Multiple Distribution lists you must format as such 'DistList1','DistList2' Example: $DistributionGroupName = 'FirmCalendar','FirmContacts';
$DistributionGroupName = "Type the Name of the distribution list here";

# Script #

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to Distribution Lists
  
$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName)
{
    Get-MailboxDatabase | Get-Mailbox -ResultSize Unlimited | Where-Object {$_.WhenCreated –ge ((Get-Date).Adddays(-1))} | Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}



Get-PSSession | Remove-PSSession
Exit

# End Scripting