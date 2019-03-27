if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Removing the MSExchDelegateListlink from an account
#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true

$domain = Read-Host "What is the name of your domain? i.e. <contonso> Do not put .com at end"
Get-ADUser -Properties msExchDelegateListlink -SearchBase "dc=$domain,dc=local" -LDAPFilter "(msExchDelegateListlink=*)" | Select-Object @{n = 'UserName'; e = {$_.userprincipalname}}, @{n = 'ListLink'; e = {$_.msExchDelegateListLink}} | Export-csv "c:\userlist.csv" –notypeinformation –AllowClobber

do {
    $UserToClean = Read-host "Type the name of the user who needs cleanup (Account name)"
    $Delegates = Get-ADUser $UserToClean -Properties msExchDelegateListlink |  Select-Object -ExpandProperty msExchDelegateListlink
    Write-Host “**************************************************************”
    Write-Host “List of Delegated accounts that are ListLinked:” $Delegates
    Write-Host “**************************************************************”
    $UserDN = Read-Host "Paste in the CN address you see above that you want to remove from msExchDelegateListlink; i.e. CN=zadd2exchange,OU=Service,DC=yourDC,DC=local"
  
    Set-ADUser $UserToClean -Remove @{msExchDelegateListLink = “$UserDN”}
  
    Write-Host “**************************************************************”
    Write-Host “If the following get-aduser cmdlet searching for ListLinks is empty, then all Delegated listlinks have been removed”
    Get-ADUser $UserToClean -Properties msExchDelegateListlink |  Select-Object -ExpandProperty msExchDelegateListlink
    Write-Host “**************************************************************”

    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')



Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting