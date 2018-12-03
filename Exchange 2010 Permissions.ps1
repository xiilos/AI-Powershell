if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


Write-Host "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
$User = read-host "Enter Sync Service Account";

$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Quit'))



$decision = $Host.UI.PromptForChoice($message, $question, $choices, 3)



add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

# Exchange 2010 on Premise-Adding new permissions all

if ($decision -eq 0) {
Write-Host "Adding Permissions to Users"
Add-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false  
Write-Host "Applying Thtottle Policy"
New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null -CPAMaxConcurrency $NULL -CPAPercentTimeInCAS $NULL -CPAPercentTimeInMailboxRPC $null
Set-Mailbox $User -ThrottlingPolicy A2EPolicy
write-host "All Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
} 
  

# Exchange 2010 on Premise-Remove old Add2Exchange permissions
  
if ($decision -eq 1) {
Write-Host "Removing Old Permissions"
Remove-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Success....."
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "All Done"
Write-Host "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}



# Option 7: Exchange on Premise- Quit

if ($decision -eq 3) {
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
  }


# End Scripting