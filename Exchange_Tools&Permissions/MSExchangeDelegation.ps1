if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Removing the MSExchDelegateListlink from an account
#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Log Path
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}


#Login to AD
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
Do {
    $UserCredential = Get-Credential
    Enter-PSSession -ComputerName $Exchangename –Credential $UserCredential -ErrorAction SilentlyContinue -ErrorVariable LoginError;
    If ($LoginError) { 
        Write-Warning -Message "Username or Password is Incorrect!"
        Write-Host "Trying Again in 2 Seconds....."
        Start-Sleep -S 2
    }
} Until (-not($LoginError))

Import-Module ActiveDirectory

$Domain = Read-Host "What is the Name of your Domain? i.e. <DidItBetter> Do not put .com at end"
Get-ADUser -Properties msExchDelegateListlink -SearchBase "dc=$domain,dc=local" -LDAPFilter "(msExchDelegateListlink=*)" | Select-Object @{n = 'UserName'; e = { $_.userprincipalname } }, @{n = 'ListLink'; e = { $_.msExchDelegateListLink } } | Export-csv "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv"

Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv"

Do {
    $UserToClean = Read-host "Type the name of the user who needs cleanup (Account name)"
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