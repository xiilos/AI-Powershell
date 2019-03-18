if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


#Goal
# Auto Assign Permissions for Add2Exchange using Windows Creds Manager

# Start of Automated Scripting #


# Option 1: Office 365


    $error.clear()
    Import-Module "MSonline" -ErrorAction SilentlyContinue
    Import-Module MSOnline
    
    Write-Host "Sign in to Office365 as Tenant Admin"
    $Cred = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    Import-Module MSOnline
    Connect-MsolService -Credential $Cred -ErrorAction Inquire
    
    $User = read-host "Service Account Name";
    $DistributionGroupName = read-host "Dist List Name";
    
    
    # Office 365-Adding to Dist. List
    
    
    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
    ForEach ($Member in $DistributionGroupName)
    {
    Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
    }
  
    
    Get-PSSession | Remove-PSSession
    Exit
      
    
    

# End Scripting