if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Service Stop #
Get-Service -ComputerName "TYPE COMPUTER NAME HERE" -Name "Add2Exchange Service" | Stop-Service -Verbose -ErrorAction Stop
Start-Sleep -s 30
Get-Service -ComputerName "TYPE COMPUTER NAME HERE" -Name "Add2Exchange Agent" | Stop-Service -Verbose
Start-Sleep -s 10

# Script #
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true
            
#Variables
# Fill Out Dynamic and Statis Distribution Groups Below

$DynamicDG = @("Dynamic DL HERE", "Dynamic DL HERE")
$StaticDG = @("Static DL HERE", "Static DL HERE")

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


#Start Service#
# Service Stop #
Get-Service -ComputerName "TYPE COMPUTER NAME HERE" -Name "Add2Exchange Service" | Start-Service -Verbose


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting