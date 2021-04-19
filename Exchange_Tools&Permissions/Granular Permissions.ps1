if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Variables:
$ServiceAccount = "Zadd2Exchange"
$Groups = Get-Content -Path "C:\zlibrary\Groupnames.txt"
$UserMailbox = ForEach ($Group in $Groups) {Get-DistributionGroupMember $Group | Select-Object -Property alias}
$Identity = ':\Contacts'
$TOPAccessRights = "FolderVisible"
$AccessRights = "Owner"


# Script #

ForEach ($Group in $Groups) {
    #Top Level
    Get-DistributionGroupMember $Group | Add-MailboxFolderPermission -User $ServiceAccount -AccessRights $TOPAccessRights

    #Target Folder
    Get-DistributionGroupMember $Group | Add-MailboxFolderPermission -Identity $usermailbox + $Identity -User $ServiceAccount -AccessRights $AccessRights
    
  
  }




  


#To the top of the mailbox
Add-MailboxFolderPermission "Awest" -User zAdd2Exchange -AccessRight FolderVisible

#To the folder
Add-MailboxFolderPermission "Awest:\Contacts" -User zAdd2Exchange -AccessRight FolderVisible

#To a subfolder
Add-MailboxFolderPermission "Awest:\Contacts\zFirm Contacts" -User zAdd2Exchange -AccessRight Owner






#Remove Permissions
Remove-MailboxFolderPermission  "<Identity>:\calendar"  -User zAdd2Exchange






#NOTES



New-Object -TypeName psobject -Property @{
    Name                     = $group.alias
    identity       = $group.$identity
    
}





Get-DistributionGroup | Select-Object DisplayName,@{n="Members";e={((Get-DistributionGroupMember -Identity $_.$identity.ToString()).displayname) -join ","}}





ForEach ($Group in $Groups) {Get-DistributionGroupMember $Group | Select-Object -ExpandProperty Name}


ForEach ($Group in $Groups) {
    Get-Mailbox -Resultsize Unlimited | Add-MailboxFolderPermission -User zadd2exchange -AccessRights $TOPAccessRights

    Get-Mailbox -Resultsize Unlimited | Add-MailboxFolderPermission -Identity $Identity -User zAdd2Exchange -AccessRights $AccessRights
  
  }




foreach($user in $users){
    Add-MailboxFolderPermission -Identity "$($user.UserPrincipalName):\Calendar" -User "my group name" -AccessRights Reviewer
}





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting