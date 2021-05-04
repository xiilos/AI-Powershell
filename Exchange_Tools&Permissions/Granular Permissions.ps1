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
$Membername = ForEach ($Group in $Groups) {Get-Distributiongroupmember $Group} #Gets the members of all groups
$Identity1 = ':\Calendar' #Default Option already filled in
$Identity2 = ':\Contacts' #Default Option already filled in
$Identity3 = ':\Contacts\zFirmContacts' #Option will be a custom feild filled in via .txt
$TOPAccessRights = "FolderVisible"
$AccessRightsFolder = "Owner"


# Script #


ForEach ($Member in $Membername) {Add-MailboxFolderPermission $Member.name -User $ServiceAccount -AccessRights $TOPAccessRights} #To top of Mailbox


ForEach ($Member in $Membername) {Add-MailboxFolderPermission ($Member.name + $Identity2) -User $ServiceAccount -AccessRights $AccessRightsFolder} #To Folder




#To Remove Rights

ForEach ($Member in $Membername) {Remove-MailboxFolderPermission $Member.name -User $ServiceAccount -confirm:$false} #From the Top

ForEach ($Member in $Membername) {Remove-MailboxFolderPermission ($Member.name + $Identity2) -User $ServiceAccount -confirm:$false} #From the Folder












#NOTES

ForEach ($Group in $Groups) {
    #Top Level
    Get-DistributionGroupMember $Group | Add-MailboxFolderPermission -User $ServiceAccount -AccessRights $TOPAccessRights

    #Target Folder
    Get-DistributionGroupMember $Group | Add-MailboxFolderPermission -Identity $usermailbox + $Identity -User $ServiceAccount -AccessRights $AccessRightsFolder
    
  
  }



  $Identity1
  $Identity2
  $Identity3
  


#To the top of the mailbox
Add-MailboxFolderPermission "Awest" -User zAdd2Exchange -AccessRight FolderVisible

#To the folder
Add-MailboxFolderPermission "Awest:\Contacts" -User zAdd2Exchange -AccessRight FolderVisible

#To a subfolder
Add-MailboxFolderPermission "Awest:\Contacts\zFirm Contacts" -User zAdd2Exchange -AccessRight Owner






#Remove Permissions
Remove-MailboxFolderPermission  "<Identity>:\calendar"  -User zAdd2Exchange

New-Object -TypeName psobject -Property @{
    Name                     = $group.alias
    identity       = $group.$identity
    
}





Get-DistributionGroup | Select-Object DisplayName,@{n="Members";e={((Get-DistributionGroupMember -Identity $_.$identity.ToString()).displayname) -join ","}}





ForEach ($Group in $Groups) {Get-DistributionGroupMember $Group | Select-Object -ExpandProperty Name}


ForEach ($Group in $Groups) {
    Get-Mailbox -Resultsize Unlimited | Add-MailboxFolderPermission -User zadd2exchange -AccessRights $TOPAccessRights

    Get-Mailbox -Resultsize Unlimited | Add-MailboxFolderPermission -Identity $Identity -User zAdd2Exchange -AccessRights $AccessRightsFolder
  
  }




foreach($user in $users){
    Add-MailboxFolderPermission -Identity "$($user.UserPrincipalName):\Calendar" -User "my group name" -AccessRights Reviewer
}





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting