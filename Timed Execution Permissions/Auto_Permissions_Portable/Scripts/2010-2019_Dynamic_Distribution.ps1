if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Script #

#Variables
$Exchangename = Get-Content ".\Add2Exchange Creds\Exchange_Server_Name.txt"
$Username = Get-Content ".\Add2Exchange Creds\Exchange_Server_Admin.txt"
$Password = Get-Content ".\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring
$DynamicDG1 = Get-Content ".\Add2Exchange Creds\Dynamic_Name.txt"
$StaticDG1 = Get-Content ".\Add2Exchange Creds\Static_Name.txt"

Try{

$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction Inquire
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true

#Timed Execution Permissions to Dynamic Distribution Lists

$DynamicDG = @($DynamicDG1)
$StaticDG = @($StaticDG1)

for ($i = 0; $i -lt $DynamicDG.Length; $i++) {
    ### Get Dynamic Group members
    $members = Get-DynamicDistributionGroup $DynamicDG[$i]
    $recipients = Get-Recipient -RecipientPreviewFilter $members.RecipientFilter
    $constrainedRecipients = @()
    foreach ($recipient in $recipients) {
        if ( $recipient.Identity.isDescendantOf( $members.RecipientContainer ) ) {
            $constrainedRecipients += $recipient
        }
    }
    #### Members are stored in $constrainedRecipients

    ### Remove the Distribution Group
    ##Changes SID for the group if done this way -> Remove-DistributionGroup -Identity $StaticDG -Confirm:$false
    $removeMembers = Get-DistributionGroupMember -Identity $StaticDG[$i]
    foreach ($remMember in $removeMembers) {
        Remove-DistributionGroupMember -Identity $StaticDG[$i] -Member $remMember -Confirm:$false
    }

    ###Create DistributionGroup members, same as the Dynamic Group
    ##New-DistributionGroup -Name $StaticDG -Type "Distribution"
    foreach ($Entry in $constrainedRecipients) {
        Add-DistributionGroupMember -Identity $StaticDG[$i] -Member $Entry
    }
}

}
  
Catch {

  Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
  Get-PSSession | Remove-PSSession
  Exit
}

  Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to a Dynamic Distribution List On-Premise Succesfully."

Get-PSSession | Remove-PSSession
Exit

# End Scripting