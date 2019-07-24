if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Variables

$Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchangename.txt"
$Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerUser.txt"
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerPass.txt" | convertto-securestring
$DynamicDG1 = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dynamic_DistributionName.txt"
$StaticDG1 = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Static_DistributionName.txt"



$Cred = New-Object -typename System.Management.Automation.PSCredential `
    -Argumentlist $Username, $Password

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred -ErrorAction SilentlyContinue
Import-PSSession $Session -DisableNameChecking
Set-ADServerSettings -ViewEntireForest $true


#Timed Execution Permissions to Dynamic Distribution Lists

$DynamicDG = @($DynamicDG1)
$StaticDG = @($StaticDG1)

for ($i = 0; $i -lt $DynamicDG.Length; $i++) {
    ### Get Dynamic Group members
    $members = Get-DynamicDistributionGroup $DynamicDG[$i]
    $Recipients = Get-Recipient -RecipientPreviewFilter $Members.RecipientFilter -OrganizationalUnit $Members.RecipientContainer
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

Get-PSSession | Remove-PSSession
Pause
Exit

# End Scripting