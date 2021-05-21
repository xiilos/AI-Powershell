if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Support Directory
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\AD_Photos"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\AD_Photos"
}

Push-Location "C:\Program Files (x86)\DidItBetterSoftware\AD_Photos"

# Script #

$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("For On-Premise AD ONLY, You Must Run this on a box with Active Directory. Click Cancel and the File will be Automatically copied to your Clipboard; Paste the file in a box with AD. 
Otherwise, Click OK to Continue.", 0, "WARNING!!", 0x1)
if ($answer -eq 2) {
    $Location = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
    Push-Location $Location
    Set-Clipboard -Path ".\Setup\Export_ADPhoto.ps1"
    Write-Host "File Copied"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}


$Title1 = 'AD Photo Export'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "Please Pick From the Following Below"
""
Write-Host "Press '1' for Exporting All Photos from all Users from Active Directory On Premise"
Write-Host "Press '2' for Exporting Photos from a Specific OU from Active Directory On Premise"
Write-Host "Press '3' for Exporting All Photos from all Users from Azure Active Directory Online"
Write-Host "Press 'Q' to Quit." -ForegroundColor Red

$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    '1' { 
        Clear-Host 
        'You Chose to Export All Photos from all Users in AD'
        Import-Module ActiveDirectory
        Do {
            $users = Get-ADUser -Filter { EmailAddress -like "*@*" } -Properties thumbnailPhoto, mail
    
            foreach ($user in $users) {
                $name = $user.mail + ".jpg"
                $user.thumbnailPhoto | Set-Content $name -Encoding byte
            }
    
            Write-Host "Finished"
    
            $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        } Until ($repeat -eq 'n')

    }

    '2' { 
        Clear-Host 
        'You Chose to Export Photos from a Specific OU'
        Import-Module ActiveDirectory
        Do {
            Write-Host "The Following Questions will Ask for the Path in which you want to search for AD Photos Example: CN=Users,DC=Diditbetter,DC=Local"
            $CN = Read-Host "Please Type in OU or CN Name"
            $DC = Read-Host "Domain Name? Ex. DC=diditbetter"
            $DC2 = Read-Host "Type in COM or Local"
    
            $users = Get-ADUser -Filter { EmailAddress -like "*@*" } -SearchBase "CN=$CN,DC=$DC,DC=$DC2" -Properties thumbnailPhoto, mail
    
            foreach ($user in $users) {
                $name = $user.mail + ".jpg"
                $user.thumbnailPhoto | Set-Content $name -Encoding byte
            }
    
            Write-Host "Finished"
    
            $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        } Until ($repeat -eq 'n')

    }

    '3' { 
        Clear-Host 
        'You Chose to Export All Photos from Azure Active Directory Online'
        Import-Module –Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        Import-Module –Name AzureAD -ErrorAction SilentlyContinue
        If ($error) {
            Write-Host "Adding Azure AD and EXO V2 module"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PSRepository -Name psgallery -InstallationPolicy Trusted
            Install-Module –Name ExchangeOnlineManagement -WarningAction "Inquire"
            Install-Module –Name AzureAD -WarningAction "Inquire"
        }
 
        Else { Write-Host 'Modules are installed' }

        Write-Host "Updating Azure AD and EXO V2 Modules Please Wait..."
        Update-Module -Name ExchangeOnlineManagement
        Import-Module –Name ExchangeOnlineManagement
        Update-Module -Name AzureAD
        Import-Module –Name AzureAD

        Write-Host "Log in as a Global Administrator"
        Connect-AzureAD

        
        
        $Name = Get-AzureADUser | Where-Object { $_.mail } | Select-Object Mail
        $Location = "C:\Program Files (x86)\DidItBetterSoftware\AD_Photos"
        $Users = Get-AzureADUser -All $true
        
        foreach ($user in $users) { Get-AzureADUserThumbnailPhoto -ObjectId $user.ObjectId -FilePath $location -FileName $user.mail }


    }

    'q' {
        Write-Host "Quitting"
        Get-PSSession | Remove-PSSession
        Exit  
    } 
}

Start-Sleep -s 2
Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\AD_Photos"


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