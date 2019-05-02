if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Support Directory
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}

Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Support"

# Script #

$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Caution... You Must Run this on a box with Active Directory. If the box you are running this on does not have Active Directory; Click Cancel and the File will be Automatically copied to your Clipboard. Otherwise, Click OK to Continue.", 0, "WARNING!!", 0x1)
if ($answer -eq 2) {
    Set-Clipboard -Path "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\MSExchangeDelegation.ps1"
    Write-Host "File Copied"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}

#Removing the MSExchDelegateListlink from an account

Import-Module ActiveDirectory

$Domain = Read-Host "What is the Name of your Domain? i.e. <DidItBetter> Do not put .com or .local at end"
Get-ADUser -Properties msExchDelegateListlink -SearchBase "dc=$domain,dc=local" -LDAPFilter "(msExchDelegateListlink=*)" | Select-Object @{n = 'UserName'; e = { $_.userprincipalname } }, @{n = 'ListLink'; e = { $_.msExchDelegateListLink } } | Export-csv "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv"

Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv"

Do {
    $UserToClean = Read-host "Type the name of the user who needs cleanup (Display name)"
    $Delegates = Get-ADUser $UserToClean -Properties msExchDelegateListlink | Select-Object -ExpandProperty msExchDelegateListlink
    Write-Host "**************************************************************"
    Write-Host “List of Delegated accounts that are ListLinked:” $Delegates
    Write-Host "**************************************************************"
    $UserDN = Read-Host "Paste in the CN address you see above that you want to remove from msExchDelegateListlink; i.e. CN=zadd2exchange,OU=Service,DC=yourDC,DC=local"
  
    Set-ADUser $UserToClean -Remove @{msExchDelegateListLink = “$UserDN” }
  
    Write-Host "**************************************************************"
    Write-Host “If the following get-aduser cmdlet searching for ListLinks is empty, then all Delegated listlinks have been removed”
    Get-ADUser $UserToClean -Properties msExchDelegateListlink | Select-Object -ExpandProperty msExchDelegateListlink
    Write-Host "**************************************************************"

    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')



Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting