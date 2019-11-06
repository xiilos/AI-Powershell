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
    $Location = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
    Push-Location $Location
    Set-Clipboard -Path ".\Setup\MSExchangeDelegation.ps1"
    Write-Host "File Copied"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}
  
#Removing the MSExchDelegateListlink from an account
  
Import-Module ActiveDirectory
Do {
    Get-ADUser -Properties msExchDelegateListlink -LDAPFilter "(msExchDelegateListlink=*)" | Select-Object @{n = 'UserName'; e = { $_.SamAccountName } }, @{n = 'ListLink'; e = { $_.msExchDelegateListLink } } | Export-csv "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv" -NoTypeInformation
  
    Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv"
  
  
    $UserDN = Read-Host "Paste in the CN address you see above that you want to remove from msExchDelegateListlink; i.e. CN=zadd2exchange,OU=Service,DC=yourDC,DC=local"
  
    $Username = Import-Csv "C:\Program Files (x86)\DidItBetterSoftware\Support\ExchangeDelegateLinkList.csv" | Select-Object Username -ExpandProperty Username
    
    $confirmation = Read-Host "Are you sure you want to remove List Links from all Users? [Y/N]"
    if ($confirmation -eq 'y') {
        Write-Host "Please Wait while we remove the List Link"
        #Removing List Link
        foreach ($username in $username) { Set-ADUser $Username -Remove @{msExchDelegateListLink = "$UserDN" } }  
        Write-Host "Finished"
    }
  
    if ($confirmation -eq 'n') {
        Write-Host "Quitting"
        Get-PSSession | Remove-PSSession
        Exit
  
    }
  
    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')
  
  
  
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
  
# End Scripting