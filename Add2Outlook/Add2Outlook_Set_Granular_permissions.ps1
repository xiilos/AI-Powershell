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
        Import-Module –Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        If ($error) {
            Write-Host "Adding EXO-V2 module"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Set-PSRepository -Name psgallery -InstallationPolicy Trusted
            Install-Module –Name ExchangeOnlineManagement -WarningAction "Inquire"
        }
     
        Else { Write-Host 'Module is installed' }

        Write-Host "Updating EXO-V2 Module Please Wait..."
        Update-Module -Name ExchangeOnlineManagement
        Import-Module –Name ExchangeOnlineManagement

        Write-Host "Sign in to Office365 as Global Admin"
    
        Connect-ExchangeOnline
  
        #Variables
        $ServiceAccount = Read-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com";
        $TOPAccessRights = "FolderVisible"
        $AccessRightsFolder = "Owner"
  
        Do {
            $Title2 = 'Office 365 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title2 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User"
            Write-Host "Press '3' for Adding Granular Permissions to a Distribution List"
            Write-Host "Press '4' for Removing Granular Permissions From a Distribution List" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
  
          
            $input2 = Read-Host "Please Make A Selection" 
            switch ($input2) {

                # Option 1: Office 365-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You Chose to Add Granular Permissions to a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)"
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to Single User"
                    Add-MailboxFolderPermission $Mailbox -User $ServiceAccount -AccessRights $TOPAccessRights
                    Add-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder
                    Write-Host "Done"
                } 
                # Option 2: Office 365-Removing Granular Permissions From a Single User
                '2' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)";
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions to Single User"
                    Remove-MailboxFolderPermission $Mailbox -User $ServiceAccount -confirm:$false
                    Remove-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -confirm:$false
                    Write-Host "Done" 
                }

                # Option 3: Office 365-Adding Granular Permissions to a Distribution List
                '3' { 
                    Clear-Host 
                    'You chose to Add Granular Permissions to a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to a Distribution List"
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission $Member.name -User $ServiceAccount -AccessRights $TOPAccessRights }
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder }
                    Write-Host "Done" 
                }

                # Option 4: Office 365-Removing Granular Permissions From a Distribution List
                '4' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions From a Distribution List"
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission $Member.name -User $ServiceAccount -confirm:$false }
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -confirm:$false }
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
                $ServiceAccountCredential = Get-Credential
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $ServiceAccountCredential -ErrorAction SilentlyContinue -ErrorVariable LoginError;
                If ($LoginError) { 
                    Write-Warning -Message "Username or Password is Incorrect!"
                    Write-Host "Trying Again in 2 Seconds....."
                    Start-Sleep -S 2
                }
            } Until (-not($LoginError))
          
      
            Import-PSSession $Session -DisableNameChecking
            Set-ADServerSettings -ViewEntireForest $true   
        }

        

        #Exchange 2010 Thottling Policy Check
        Write-Host "Checking Throttling Policy"
        $ThrottlePolicy = Get-ThrottlingPolicy -identity A2EPolicy -ErrorAction SilentlyContinue ;
        If ($ThrottlePolicy = $ThrottlePolicy) {
            Write-Host "Throttling Policy Exists... Adding $ServiceAccount"
            Set-Mailbox $ServiceAccount -ThrottlingPolicy A2EPolicy -WarningAction SilentlyContinue ; Write-Host "$ServiceAccount is Now a Member of the A2EPolicy"
        }

        Else {
            Write-Host "Creating New Throttling Policy and Adding $ServiceAccount"
            New-ThrottlingPolicy A2EPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
            Set-Mailbox $ServiceAccount -ThrottlingPolicy A2EPolicy
            Write-Host "$ServiceAccount is Now a Member of the A2EPolicy"
        } 
  

        #Variables
        $ServiceAccount = Read-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com";
        $TOPAccessRights = "FolderVisible"
        $AccessRightsFolder = "Owner"


        Do {

            $Title3 = 'Exchange 2010 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title3 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User"
            Write-Host "Press '3' for Adding Granular Permissions to a Distribution List"
            Write-Host "Press '4' for Removing Granular Permissions From a Distribution List" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
  
          
            $input3 = Read-Host "Please Make A Selection" 
            switch ($input3) {

                # Option 1: Exchange 2010-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You Chose to Add Granular Permissions to a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)"
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to Single User"
                    Add-MailboxFolderPermission $Mailbox -User $ServiceAccount -AccessRights $TOPAccessRights
                    Add-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder
                    Write-Host "Done"
                } 
                # Option 2: Exchange 2010-Removing Granular Permissions From a Single User
                '2' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)";
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions to Single User"
                    Remove-MailboxFolderPermission $Mailbox -User $ServiceAccount -confirm:$false
                    Remove-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -confirm:$false
                    Write-Host "Done" 
                }

                # Option 3: Exchange 2010-Adding Granular Permissions to a Distribution List
                '3' { 
                    Clear-Host 
                    'You chose to Add Granular Permissions to a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to a Distribution List"
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission $Member.name -User $ServiceAccount -AccessRights $TOPAccessRights }
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder }
                    Write-Host "Done" 
                }

                # Option 4: Exchange 2010-Removing Granular Permissions From a Distribution List
                '4' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions From a Distribution List"
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission $Member.name -User $ServiceAccount -confirm:$false }
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -confirm:$false }
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
                $ServiceAccountCredential = Get-Credential
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $ServiceAccountCredential -ErrorAction SilentlyContinue -ErrorVariable LoginError;
                If ($LoginError) { 
                    Write-Warning -Message "Username or Password is Incorrect!"
                    Write-Host "Trying Again in 2 Seconds....."
                    Start-Sleep -S 2
                }
            } Until (-not($LoginError))
      
            Import-PSSession $Session -DisableNameChecking
            Set-ADServerSettings -ViewEntireForest $true   
        }


        #Variables
        $ServiceAccount = Read-Host "Enter Sync Service Account name (Display Name) Example: zAdd2Exchange or zAdd2Exchange@domain.com";
        $TOPAccessRights = "FolderVisible"
        $AccessRightsFolder = "Owner"

        #Exchange 2013-2019 Thottling Policy Check
        Write-Host "Checking Throttling Policy"
        $ThrottlePolicy = Get-ThrottlingPolicy -identity A2EPolicy -ErrorAction SilentlyContinue ;
        If ($ThrottlePolicy = $ThrottlePolicy) {
            Write-Host "Throttling Policy Exists... Adding $ServiceAccount"
            Set-ThrottlingPolicyAssociation $ServiceAccount -ThrottlingPolicy A2EPolicy -WarningAction SilentlyContinue ; Write-Host "$ServiceAccount is Now a Member of the A2EPolicy"
        }

        Else {
            Write-Host "Creating New Throttling Policy and Adding $ServiceAccount"
            New-ThrottlingPolicy -Name A2EPolicy -RCAMaxConcurrency Unlimited -EWSMaxConcurrency Unlimited
            Set-ThrottlingPolicyAssociation $ServiceAccount -ThrottlingPolicy A2EPolicy
            Write-Host "$ServiceAccount is Now a Member of the A2EPolicy"
        }              
   
        Do {

            $Title4 = 'Exchange 2013-2019 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title4 ================" 
            ""     
            Write-Host "Press '1' for Adding Granular Permissions to a Single User" 
            Write-Host "Press '2' for Removing Granular Permissions From a Single User"
            Write-Host "Press '3' for Adding Granular Permissions to a Distribution List"
            Write-Host "Press '4' for Removing Granular Permissions From a Distribution List" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
  
          
            $input4 = Read-Host "Please Make A Selection" 
            switch ($input4) {

                # Option 1: Exchange 2013-2019-Adding Granular Permissions to a Single User
                '1' { 
                    Clear-Host 
                    'You Chose to Add Granular Permissions to a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)"
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to Single User"
                    Add-MailboxFolderPermission $Mailbox -User $ServiceAccount -AccessRights $TOPAccessRights
                    Add-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder
                    Write-Host "Done"
                } 
                # Option 2: Exchange 2013-2019-Removing Granular Permissions From a Single User
                '2' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Single User'
                    $Mailbox = Read-Host "(Enter user Email Address)";
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions to Single User"
                    Remove-MailboxFolderPermission $Mailbox -User $ServiceAccount -confirm:$false
                    Remove-MailboxFolderPermission ($Mailbox + $Identity) -User $ServiceAccount -confirm:$false
                    Write-Host "Done" 
                }

                # Option 3: Exchange 2013-2019-Adding Granular Permissions to a Distribution List
                '3' { 
                    Clear-Host 
                    'You chose to Add Granular Permissions to a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Adding Granular Permissions to a Distribution List"
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission $Member.name -User $ServiceAccount -AccessRights $TOPAccessRights }
                    ForEach ($Member in $Membername) { Add-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -AccessRights $AccessRightsFolder }
                    Write-Host "Done" 
                }

                # Option 4: Exchange 2013-2019-Removing Granular Permissions From a Distribution List
                '4' { 
                    Clear-Host 
                    'You chose to Remove Granular Permissions From a Distribution List'
                    $Groups = Read-Host "Enter distribution list name (Display Name)";
                    $Membername = ForEach ($Group in $Groups) { Get-Distributiongroupmember $Group }
                    $Identity = Read-Host "Enter the path of the destination folder :\Calendar or :\Contact (Example :\Contacts\zFirm Contacts)"
                    Write-Host "Removing Granular Permissions From a Distribution List"
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission $Member.name -User $ServiceAccount -confirm:$false }
                    ForEach ($Member in $Membername) { Remove-MailboxFolderPermission ($Member.name + $Identity) -User $ServiceAccount -confirm:$false }
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