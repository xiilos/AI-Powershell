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

$Title1 = 'Add2Outlook Granular Permissions Menu'

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
        Import-Module "MSonline" -ErrorAction SilentlyContinue
        If ($error) {
            Write-Host "Adding Azure MSonline module"
            Set-PSRepository -Name psgallery -InstallationPolicy Trusted
            Install-Module MSonline -Confirm:$false -WarningAction "Inquire"
        } 
        Else { Write-Host 'Module is installed' }


        Write-Host "Sign in to Office365 as Tenant Admin"
        Do {
            $Cred = Get-Credential
            Connect-MsolService -Credential $Cred
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection -ErrorAction SilentlyContinue -ErrorVariable LoginError;
            If ($LoginError) { 
                Write-Warning -Message "Username or Password is Incorrect!"
                Write-Host "Trying Again in 2 Seconds....."
                Start-Sleep -S 2
            }
        } Until (-not($LoginError))
        
        Import-PSSession $Session -DisableNameChecking
        Import-Module MSOnline
    
        $User = Read-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com";
    
        Do {
            $Title2 = 'Office 365 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title2 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
    
            
            $input2 = Read-Host "Please Make A Selection" 
            switch ($input2) {

                # Option 1: Office 365-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You Chose to Add Granular Permissions to a Single User'
                    $Mailbox = read-host "(Enter user Email Address)"
                    $Identity = read-host "Enter user Email Address followed by :\Calendar or :\Contact (Example: Tom@diditbetter.com:\Contacts)"
                    $AccessRights = read-host "Enter Permissions level (Owner, Editor, none)"
                    Write-Host "Adding Add2Outlook Granular Permissions to Single User"
                    Add-MailboxPermission -Identity $Mailbox -User $User -AccessRights 'readpermission'
                    Add-MailboxFolderPermission -Identity $Identity -User $User -AccessRights $AccessRights
                    Write-Host "Done"
                } 
                # Option 2: Office 365-Removing Granular Permissions From a Single User
                '2' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $User = read-host "Enter Sync Service Account (Display Name)";
                    $Identity = read-host "Enter user Email Address"
                    Write-Host "Removing Add2Outlook Granular Permissions to Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done" 
                }
                
                # Option Q: Office 365-Quit
                'q' { 
                    Write-Host "Quitting"
                    Get-PSSession | Remove-PSSession
                    Exit  
                } 
            }
            $repeat = Read-Host 'Return To The Main Menu? [Y/N]'
        } Until ($repeat -eq 'n')
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
  
        Write-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com"
        $User = Read-Host "Enter Sync Service Account";
  
        #Exchange 2010 Thottling Policy Check
        Write-Host "Checking Throttling Policy"
        $ThrottlePolicy = Get-ThrottlingPolicy -identity A2EPolicy -ErrorAction SilentlyContinue ;
        If ($ThrottlePolicy = $ThrottlePolicy) {
            Write-Host "Throttling Policy Exists... Adding $User"
            Set-Mailbox $User -ThrottlingPolicy A2EPolicy -WarningAction SilentlyContinue ; Write-Host "$User is Now a Member of the A2EPolicy"
        }

        Else {
            Write-Host "Creating New Throttling Policy and Adding $User"
            New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
            Set-Mailbox $User -ThrottlingPolicy A2EPolicy
            Write-Host "$User is Now a Member of the A2EPolicy"
        } 
    
        Do {

            $Title3 = 'Exchange 2010 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title3 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red

            $input3 = Read-Host "Please Make A Selection" 
            switch ($input3) {

                # Option 1: Exchange 2010 on Premise-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You chose to Add Granular Permissions to a Single User'
                    $Mailbox = read-host "(Enter user Email Address)"
                    $Identity = read-host "Enter user Email Address followed by :\Calendar or :\Contact (Example: Tom@diditbetter.com:\Contacts)"
                    $AccessRights = read-host "Enter Permissions level (Owner, Editor, none)"
                    Write-Host "Adding Add2Outlook Granular Permissions to Single User"
                    Add-MailboxPermission -Identity $Mailbox -User $User  -AccessRights 'readpermission'
                    Add-MailboxFolderPermission -Identity $Identity -User $User -AccessRights $AccessRights
                    Write-Host "Done"
                }
                # Option 2: Exchange 2010 on Premise-Removing Granular Permissions From a Single User
                '2' {
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $Identity = read-host "Enter user Email Address"
                    Write-Host "Removing Add2Outlook Granular Permissions From a Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }
                
                # Option Q: Exchange 2010-Quit
                'q' { 
                    Write-Host "Quitting"
                    Get-PSSession | Remove-PSSession
                    Exit 
                }
            }
            $repeat = Read-Host 'Return to the Main Menu? [Y/N]'
        } Until ($repeat -eq 'n')
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


        Write-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com"
        $User = Read-Host "Enter Sync Service Account";

        #Exchange 2013-2019 Thottling Policy Check
        Write-Host "Checking Throttling Policy"
        $ThrottlePolicy = Get-ThrottlingPolicy -identity A2EPolicy -ErrorAction SilentlyContinue ;
        If ($ThrottlePolicy = $ThrottlePolicy) {
            Write-Host "Throttling Policy Exists... Adding $User"
            Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy -WarningAction SilentlyContinue ; Write-Host "$User is Now a Member of the A2EPolicy"
        }

        Else {
            Write-Host "Creating New Throttling Policy and Adding $User"
            New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
            Set-ThrottlingPolicyAssociation $User -ThrottlingPolicy A2EPolicy
            Write-Host "$User is Now a Member of the A2EPolicy"
        }              
     
        Do {

            $Title4 = 'Exchange 2013-2019 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title4 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red

            $input3 = Read-Host "Please Make A Selection" 
            switch ($input3) {

                # Option 1: Exchange 2013-2019 on Premise-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You chose to Add Granular Permissions to a Single User'
                    $Mailbox = read-host "(Enter user Email Address)"
                    $Identity = read-host "Enter user Email Address followed by :\Calendar or :\Contact (Example: Tom@diditbetter.com:\Contacts)"
                    $AccessRights = read-host "Enter Permissions level (Owner, Editor, none)"
                    Write-Host "Adding Add2Outlook Granular Permissions to Single User"
                    Add-MailboxPermission -Identity $Mailbox -User $User  -AccessRights 'readpermission'
                    Add-MailboxFolderPermission -Identity $Identity -User $User -AccessRights $AccessRights
                    Write-Host "Done"
                }
                # Option 2: Exchange 2013-2019 on Premise-Removing Granular Permissions From a Single User
                '2' {
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $Identity = read-host "Enter user Email Address"
                    Write-Host "Removing Add2Outlook Granular Permissions From a Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }

                # Option Q: Exchange 2013-2019-Quit
                'q' { 
                    Write-Host "Quitting"
                    Get-PSSession | Remove-PSSession
                    Exit 
                }
            }
            $repeat = Read-Host 'Return to the Main Menu? [Y/N]'
        } Until ($repeat -eq 'n')
    }       

}#Origin LogonMethod End


#Quit All
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Start-Sleep -s 1
Exit    
    

# End Scripting