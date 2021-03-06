if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}

Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Permissions.txt" -Append

# Script #

$Title1 = 'Exchange Shell'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "How Are We Logging In?"
""
Write-Host "Press '1' for Office 365"
Write-Host "Press '2' for Exchange 2010" 
Write-Host "Press '3' for Exchange 2013-2019" 
Write-Host "Press 'Q' to Quit." -ForegroundColor Red


#Login Method
 
$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    #Office 365--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '1' { 
        Clear-Host 
        'You chose Office 365'
        $error.clear()
        Import-Module �Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        If ($error) {
            Write-Host "Adding EXO-V2 module"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PSRepository -Name psgallery -InstallationPolicy Trusted
            Install-Module �Name ExchangeOnlineManagement -WarningAction "Inquire"
        }
         
        Else { Write-Host 'Module is installed' }

        Write-Host "Updating EXO-V2 Module Please Wait..."
        Update-Module -Name ExchangeOnlineManagement
        Import-Module �Name ExchangeOnlineManagement

        Write-Host "Sign in to Office365 as Global Admin"
        
        Connect-ExchangeOnline
    
    }
            
        
    #Exchange2010--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '2' {
        Clear-Host 
        'You chose Exchange 2010'
        $confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
        if ($confirmation -eq 'y') {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
            Set-ADServerSettings -ViewEntireForest $true
        }
    
        if ($confirmation -eq 'n') {
            Write-Warning -Message "Before Continuing, please remote into your Exchange server.
            Open Powershell as administrator and Type: *Enable-PSRemoting* without the stars and hit enter.
            Once Done, click Enter to Continue"
            Pause
            
            $Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
            Do {
                $UserCredential = Get-Credential
                If (!$UserCredential) { Exit }
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction SilentlyContinue -ErrorVariable LoginError;
                If ($LoginError) { 
                    Write-Warning -Message "Username or Password is Incorrect!"
                    Write-Host "Trying Again in 2 Seconds....."
                    Start-Sleep -S 2
                }
            } Until (-not($LoginError))
            
        
            Import-PSSession $Session -DisableNameChecking
            Set-ADServerSettings -ViewEntireForest $true   
        }
  
       
    }
    #Exchange2013-2019--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '3' { 
        Clear-Host 
        'You chose Exchange 2013-2019'
        $confirmation = Read-Host "Are you on the Exchange Server? [Y/N]"
        if ($confirmation -eq 'y') {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
            Set-ADServerSettings -ViewEntireForest $true
        }

        if ($confirmation -eq 'n') {
            Write-Warning -Message "Before Continuing, please remote into your Exchange server.
            Open Powershell as administrator and Type: *Enable-PSRemoting* without the stars and hit enter.
            Once Done, click Enter to Continue"
            Pause
        
            $Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
            Do {
                $UserCredential = Get-Credential
                If (!$UserCredential) { Exit }
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction SilentlyContinue -ErrorVariable LoginError;
                If ($LoginError) { 
                    Write-Warning -Message "Username or Password is Incorrect!"
                    Write-Host "Trying Again in 2 Seconds....."
                    Start-Sleep -S 2
                }
            } Until (-not($LoginError))
        
            Import-PSSession $Session -DisableNameChecking
            Set-ADServerSettings -ViewEntireForest $true   
        }
    }       


    #Quit All
    'q' { 
        Clear-Host 
        Write-Host "ttyl"
        Get-PSSession | Remove-PSSession
        Start-Sleep -s 1
        Exit
    }
}#Origin LogonMethod End




# End Scripting