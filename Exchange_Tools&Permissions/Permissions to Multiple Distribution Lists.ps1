if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Add Permissions

$Groups = Get-Content -Path "C:\zlibrary\Groupnames.txt"


ForEach ($Group in $Groups) {
  Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User zadd2exchange -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false

}

#Remove Permissions

$Groups = Get-Content -Path "C:\zlibrary\Groupnames.txt"


ForEach ($Group in $Groups) {
    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User zAdd2Exchange -accessrights FullAccess -verbose -confirm:$false

}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting