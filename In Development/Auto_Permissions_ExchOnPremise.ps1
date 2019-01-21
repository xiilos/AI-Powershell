if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


#Goal
# Auto Assign Permissions for Add2Exchange using Windows Creds Manager

# Start of Automated Scripting #

# Exchange On Premise

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://Exchange2016.mbx.local/PowerShell/ -Authentication credssp -Credential (Get-StoredCredential -Target "Exchange2016.mbx.local") -ErrorAction Inquire
    Import-PSSession $Session -DisableNameChecking
    Set-ADServerSettings -ViewEntireForest $true   
    
    
    # Exchange on Premise-Adding Permissions to dist. list
    
    Get-DistributionGroupMember "zcalendars"
    ForEach ($Member in "zcalendars")
    {
    Add-MailboxPermission -Identity $Member.name -User 'zadd2exchange' -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
    }
    
    
    Get-PSSession | Remove-PSSession
    Exit
      
    
    

# End Scripting