if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Service_Account_Name.txt"
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Admin_Pass.txt" | convertto-securestring
$DynamicGroupName = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dynamic_Name.txt"
$StaticGroupName = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Static_Name.txt"

Try {

$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

#Timed Execution Permissions to Distribution Lists
for ($i = 0; $i -lt $DynamicGroupName.Length; $i++) {
    ### Get Dynamic Group members
    $members = Get-DynamicDistributionGroup $DynamicGroupName[$i]
    $recipients = Get-Recipient -RecipientPreviewFilter $members.RecipientFilter
    $constrainedRecipients = @()
    foreach ($recipient in $recipients) {
        if ( $recipient.Identity.isDescendantOf( $members.RecipientContainer ) ) {
            $constrainedRecipients += $recipient
        }
    }
    #### Members are stored in $constrainedRecipients

    ### Remove the Distribution Group
    ##Changes SID for the group if done this way -> Remove-DistributionGroup -Identity $StaticGroupName -Confirm:$false
    $removeMembers = Get-DistributionGroupMember -Identity $StaticGroupName[$i]
    foreach ($remMember in $removeMembers) {
        Remove-DistributionGroupMember -Identity $StaticGroupName[$i] -Member $remMember -Confirm:$false
    }

    ###Create DistributionGroup members, same as the Dynamic Group
    ##New-DistributionGroup -Name $StaticGroupName -Type "Distribution"
    foreach ($Entry in $constrainedRecipients) {
        Add-DistributionGroupMember -Identity $StaticGroupName[$i] -Member $Entry
    }
}


}

Catch {
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Get-PSSession | Remove-PSSession
    Exit
  }
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to a Dynamic Distribution List On Office 365 Succesfully."
  
  Get-PSSession | Remove-PSSession
  Exit

# End Scripting