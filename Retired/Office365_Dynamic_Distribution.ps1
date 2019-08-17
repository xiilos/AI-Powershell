if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Variables #
$TenentUsername = Get-StoredCredential -target 'Office365_Tenent' -ascredentialobject | Select-Object Username -ExpandProperty Username
$TenentPassword = Get-StoredCredential -target 'Office365_Tenent' | Select-Object password -ExpandProperty password
$DynamicDG = Get-StoredCredential -target 'Dynamic_Distribution_Group' -ascredentialobject | Select-Object username -ExpandProperty username
$StaticDG = Get-StoredCredential -target 'Static_Distribution_Group' -ascredentialobject | Select-Object username -ExpandProperty username

# Script #

$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $TenentUsername, $TenentPassword

Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Import-Module MSOnline

#Timed Execution Permissions to Dynamic Distribution Lists

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


Get-PSSession | Remove-PSSession
Exit

# End Scripting