if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

#Support Directory
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"
}

        #Check for MS Online Module
        Write-Host "Checking for Exhange Online Module"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        IF (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
            Write-Host "Exchange Online Module Exists"
    
            $InstalledEXOv2 = ((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()

            $LatestEXOv2 = (Find-Module -Name ExchangeOnlineManagement).Version.ToString()
    
            [PSCustomObject]@{
                Match = If ($InstalledEXOv2 -eq $LatestEXOv2) { Write-Host "You are on the latest Version" } 
        
                Else {
                    Write-Host "Upgrading Modules..."
                    Update-Module -Name ExchangeOnlineManagement -Force
                    Write-Host "Success"
                }

            }


        } 
        Else {
            Write-Host "Module Does Not Exist"
            Write-Host "Downloading Exchange Online Management..."
            Install-Module –Name ExchangeOnlineManagement -Force
            Write-Host "Success"
        }


# Script #

Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"

Import-Module –Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
Import-Module –Name AzureAD -ErrorAction SilentlyContinue
If ($error) {
    Write-Host "Adding Azure AD module"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-PSRepository -Name psgallery -InstallationPolicy Trusted
    Install-Module –Name AzureAD -WarningAction "Inquire"
}
 
Else { Write-Host 'Modules are installed' }

Write-Host "Updating Azure AD Modules Please Wait..."
Update-Module -Name AzureAD
Import-Module –Name AzureAD

Connect-AzureAD


$Name = Get-AzureADUser | Where-Object {$_.mail} | Select-Object Mail
$Location = "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"
$Users = Get-AzureADUser -All $true



foreach ($user in $users) {Get-AzureADUserThumbnailPhoto -ObjectId $user.ObjectId -FilePath $location -FileName $user.mail}


Start-Sleep -s 2
Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"

$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Finished Exporting Azure AD Photos", 0, "Complete", 0x1)
if ($answer -eq 2) {
    Break
}



Write-Host "ttyl"
Disconnect-AzureAD
Get-PSSession | Remove-PSSession
Exit

# End Scripting