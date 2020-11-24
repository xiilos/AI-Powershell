if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$ServiceAccount = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
$Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Service_Account_Name.txt"
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Admin_Pass.txt" | convertto-securestring

Try {

$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

#Timed Execution Permissions to All Users
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $ServiceAccount -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
}


Catch {
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Get-PSSession | Remove-PSSession
    Exit
  }
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to All Users on Office 365 Succesfully."
  
  Get-PSSession | Remove-PSSession
  Exit

# End Scripting