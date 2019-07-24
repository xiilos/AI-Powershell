if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

$DistributionGroupName = @("zcalendars", "zcontacts")


$DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
ForEach ($Member in $DistributionGroupName) {
    Add-MailboxPermission -Identity $Member.name -User 'zadd2exchange' -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
}






foreach($group in (get-content c:\grouplist.txt)){
    Remove-DistributionGroupMember -Identity $group -Member $username -Confirm:$false -Verbose
}





#Checking Source Distribution List Name


$TestPath = ".\DistributionList.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    Write-Host "Distribution List Name File Exists..."
    Write-Host "Current Content of File:" -ForegroundColor Green
    Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\DistributionList.txt"
    ""
    $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
    if ($confirmation -eq 'N') {
        Write-Host "Resuming"
    }
Do {
    if ($confirmation -eq 'Y') {
        Read-Host "Type in the Distribution List Name if adding permissions to a list. Leave Blank for None. Press Enter when Finished." | Add-content "c:\DistributionList.txt"
    }
        
    } until ($confirmation -eq 'N')
}

Else {
    Read-Host "If Adding Permissions to a Distribution List Type in the Distribution List Name. Leave Blank for None. Press Enter when Finished.
    For Multiple Entries Example: zFirmContacts, zFirmCalendar
    Note: Make sure to use qoutes around each group seperated by a comma." | out-file ".\DistributionList.txt"
}





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting