if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Support Directory
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"
}

Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Support\AD_Photos"

# Script #

$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Caution... You Must Run this on a box with Active Directory. If the box you are running this on does not have Active Directory; Cancel this Window and copy the file to an appropriate box", 0, "WARNING!!", 0x1)
if ($answer -eq 2) { Break }

Import-Module ActiveDirectory

$Title1 = 'AD Photo Export'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "Please Pick From the Following Below"
""
Write-Host "Press '1' for Exporting All Photos from all Users in AD"
Write-Host "Press '2' for Exporting Photos from a Specific OU"  
Write-Host "Press 'Q' to Quit." -ForegroundColor Red

$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    '1' { 
        Clear-Host 
        'You Chose to Export All Photos from all Users in AD'
        Do {
            $users = Get-ADUser -Filter { EmailAddress -like "*@*" } -Properties thumbnailPhoto
    
            foreach ($user in $users) {
                $name = $user.UserPrincipalName + ".jpg"
                $user.thumbnailPhoto | Set-Content $name -Encoding byte
            }
    
            Write-Host "Finished"
    
            $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        } Until ($repeat -eq 'n')

    }

    '2' { 
        Clear-Host 
        'You Chose to Export Photos from a Specific OU'
        Do {
            Write-Host "The Following Questions will Ask for the Path in which you want to search for AD Photos Example: CN=Users,DC=Diditbetter,DC=Local"
            $CN = Read-Host "Please Type in OU or CN Name"
            $DC = Read-Host "Domain Name? DC=diditbetter"
            $DC2 = Read-Host "Type in COM or Local"
    
            $users = Get-ADUser -Filter { EmailAddress -like "*@*" } -SearchBase "CN=$CN,DC=$DC,DC=$DC2" -Properties thumbnailPhoto
    
            foreach ($user in $users) {
                $name = $user.UserPrincipalName + ".jpg"
                $user.thumbnailPhoto | Set-Content $name -Encoding byte
            }
    
            Write-Host "Finished"
    
            $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        } Until ($repeat -eq 'n')

    }

    'q' { 
        Write-Host "Quitting"
        Get-PSSession | Remove-PSSession
        Exit  
    } 
}

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting