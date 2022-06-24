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

Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Permission_Results.txt" -Append

# Script #

$Title1 = 'Add2Exchange Enterprise Permissions Menu'

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

        #Variables
        $User = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
        $Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Service_Account_Name.txt"
        $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Admin_Pass.txt" | convertto-securestring

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
        
        Import-Module –Name ExchangeOnlineManagement -ErrorAction SilentlyContinue

        Write-Host "Signing in to Office365 as Exchange Admin"
        
        $Cred = New-Object -typename System.Management.Automation.PSCredential `
            -Argumentlist $Username, $Password
    
            Do {Connect-ExchangeOnline -Credential $Cred -ErrorAction SilentlyContinue -ErrorVariable ConnectError;
                If ($ConnectError)
                {Write-Host "Could not pass credentials. Trying manually..." 
                Connect-Exchangeonline}
                
                
                } Until (-not($ConnectError))
    
    
        Do {
            $Title2 = 'Office 365 Permissions Menu' 
            ""
            Clear-Host 
            Write-Host "================ $Title2 ================" 
            ""     
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Removing Permissions from All Users" 
            Write-Host "Press '3' for Removing and then Adding Permissions to All Users"
            Write-Host "Press '4' for Adding Permissions to a Distribution List"
            Write-Host "Press '5' for Removing Permissions from a Distribution List"
            Write-Host "Press '6' for Adding Permissions to a Single User"
            Write-Host "Press '7' for Removing Permissions from a Single User"
            Write-Host "Press '8' for Adding Permissions to Public Folders"
            Write-Host "Press '9' for Removing Permissions From Public Folders"
            Write-Host "Press '10' for Adding Add2Exchange Permissions to All Public Folders"
            Write-Host "Press '11' for Removing Add2Exchange Permissions From All Public Folders" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
    
            
            $input2 = Read-Host "Please Make A Selection" 
            switch ($input2) {

                # Option 1: Office 365-Adding Add2Exchange Permissions
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                    Write-Host "Adding Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
                    Write-Host "Done"
                } 
                # Option 2: Office 365-Removing Add2Exchange Permissions
                '2' { 
                    Clear-Host 
                    'You chose to Remove Permissions from All Users'
                    Write-Host "Removing Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -verbose -confirm:$false
                    Write-Host "Done" 
                } 
                # Option 3: Office 365-Remove&Add Permissions
                '3' { 
                    Clear-Host 
                    'You chose to Remove and then Add Permissions to All Users'
                    Write-Host "Removing Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -Verbose -confirm:$false
                    Write-Host "Adding Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
                    Write-Host "Done"
                } 
                # Option 4: Office 365-Adding to Dist. List
                '4' { 
                    Clear-Host 
                    'You chose to Add Permissions to a Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Adding Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                        Write-Host "Done"
                    }
                }
                # Option 5: Office 365-Remove Permissions within a dist. list
                '5' { 
                    Clear-Host 
                    'You chose to Remove Permissions From a Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Removing Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                        Write-Host "Done"
                    }
                }
                # Option 6: Office 365-Add Permissions to single user
                '6' { 
                    Clear-Host 
                    'You chose to Add Permissions to a Single User'
                    $Identity = Read-Host "Enter user Email Address"
                    Write-Host "Adding Add2Exchange Permissions to Single User"
                    Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                    Write-Host "Done"
                } 
                # Option 7: Office 365-Remove Single user permissions
                '7' { 
                    Clear-Host 
                    'You chose to Remove Permissions From a Single User'
                    $Identity = Read-Host "Enter user Email Address"
                    Write-Host "Removing Add2Exchange Permissions to Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }
                # Option 8: Office 365-Adding Permissions to Public Folders
                '8' { 
                    Clear-Host 
                    'You chose to Add Permissions to Public Folders'
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Adding Permissions to Public Folders"
                    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                        Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }
                # Option 9: Office 365-Removing Permissions From Public Folders
                '9' { 
                    Clear-Host 
                    'You chose to Remove Permissions From Public Folders'
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Removing Permissions to Public Folders"
                    Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                    Write-Host "Done"
                }

                # Option 10: Office 365-Add Permissions to All Public Folders
                '10' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Public Folders'
                    Write-Host "Adding Add2Exchange as Owner to All Public Folders"
                    Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Get-PublicFolder –Identity "\" –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
                        Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }

                # Option 11: Office 365-Removing Permissions From All Public Folders
                '11' { 
                    Clear-Host 
                    'You chose to Remove Add2Exchange Permissions from All Public Folders'
                    Write-Host "Removing Add2Exchange Owner Permissions from All Public Folders"
                    Get-PublicFolder –Identity "\" –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
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
        #Variables
        $Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Name.txt"
        $User = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
        $Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
        $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring
  
        $Cred = New-Object -typename System.Management.Automation.PSCredential `
            -Argumentlist $Username, $Password

        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred
        Import-PSSession $Session -DisableNameChecking
        Set-ADServerSettings -ViewEntireForest $true

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
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Removing Permissions from All Users" 
            Write-Host "Press '3' for Removing and then Adding Permissions to All Users"
            Write-Host "Press '4' for Adding Permissions to a Distribution List"
            Write-Host "Press '5' for Removing Permissions from a Distribution List"
            Write-Host "Press '6' for Adding Permissions to a Single User"
            Write-Host "Press '7' for Removing Permissions from a Single User"
            Write-Host "Press '8' for Adding Permissions to Public Folders"
            Write-Host "Press '9' for Removing Permissions From Public Folders"
            Write-Host "Press '10' for Adding Add2Exchange Permissions to All Public Folders"
            Write-Host "Press '11' for Removing Add2Exchange Permissions From All Public Folders"
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red

            $input3 = Read-Host "Please Make A Selection" 
            switch ($input3) {

                # Option 1: Exchange 2010 on Premise-Adding new permissions all
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                    Write-Host "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Host "Done"
                }
                # Option 2: Exchange 2010 on Premise-Remove old Add2Exchange permissions
                '2' {
                    Clear-Host 
                    'You chose to Remove Permissions to All Users'
                    Write-Host "Removing Old zAdd2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }
                # Option 3: Exchange 2010 on Premise-Remove/Add Permissions all
                '3' {
                    Clear-Host 
                    'You chose to Remove and then Add Permissions to All Users'
                    Write-Host "Removing Old zAdd2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Success....."
                    Write-Host "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Host "Checking............"
                    Write-Host "Done"
                }
                # Option 4: Exchange 2010 on Premise-Adding Permissions to dist. list
                '4' {
                    Clear-Host 
                    'You chose to Add Permissions To A Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Adding Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                        Write-Host "Done"
                    }
                }
                # Option 5: Exchange 2010 on Premise-Removing dist. list permissions
                '5' {
                    Clear-Host 
                    'You chose to Remove Permissions From A Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Removing Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                        Write-Host "Done"
                    }
                }
                # Option 6: Exchange 2010 on Premise-Adding permissions to single user
                '6' {
                    Clear-Host 
                    'You chose to Add Permissions To A Single User'
                    $Identity = Read-Host "Enter user Email Address";
                    Write-Host "Adding Add2Exchange Permissions to Single User"
                    Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                    Write-Host "Done" 
                }
                # Option 7: Exchange 2010 on Premise-Removing permissions to single user
                '7' {
                    Clear-Host 
                    'You chose to Remove Permissions To A Single User'
                    $Identity = Read-Host "Enter user Email Address"
                    Write-Host "Removing Add2Exchange Permissions to Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }
                # Option 8: Exchange 2010-Adding Permissions to Public Folders
                '8' { 
                    Clear-Host 
                    'You chose to Add Permissions to Public Folders'
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Adding Permissions to Public Folders"
                    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                        Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }
                # Option 9: Exchange 2010-Removing Permissions From Public Folders
                '9' { 
                    Clear-Host
                    'You chose to Remove Permissions From Public Folders' 
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Removing Permissions to Public Folders"
                    Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                    Write-Host "Done"
                }

                # Option 10: Exchange 2010-Add Permissions to All Public Folders
                '10' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Public Folders'
                    Write-Host "Adding Add2Exchange as Owner to All Public Folders"
                    Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Get-PublicFolder –Identity "\" –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
                        Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }

                # Option 11: Exchange 2010-Removing Permissions From All Public Folders
                '11' { 
                    Clear-Host 
                    'You chose to Remove Add2Exchange Permissions from All Public Folders'
                    Write-Host "Removing Add2Exchange Owner Permissions from All Public Folders"
                    Get-PublicFolder –Identity "\" –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
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
        #Variables
        $Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Name.txt"
        $User = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
        $Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
        $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring
          
        $Cred = New-Object -typename System.Management.Automation.PSCredential `
            -Argumentlist $Username, $Password
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred
        Import-PSSession $Session -DisableNameChecking
        Set-ADServerSettings -ViewEntireForest $true

                
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
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Removing Permissions from All Users" 
            Write-Host "Press '3' for Removing and then Adding Permissions to All Users"
            Write-Host "Press '4' for Adding Permissions to a Distribution List"
            Write-Host "Press '5' for Removing Permissions from a Distribution List"
            Write-Host "Press '6' for Adding Permissions to a Single User"
            Write-Host "Press '7' for Removing Permissions from a Single User"
            Write-Host "Press '8' for Adding Permissions to Public Folders"
            Write-Host "Press '9' for Removing Permissions From Public Folders"
            Write-Host "Press '10' for Adding Add2Exchange Permissions to All Public Folders"
            Write-Host "Press '11' for Removing Add2Exchange Permissions From All Public Folders"
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red

            
            $input4 = Read-Host "Please Make A Selection" 
            switch ($input4) {
                # Option 1: Exchange 2013-2019 on Premise-Adding new permissions all
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                    Write-Host "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Host "Done"
                }
                # Option 2: Exchange 2013-2019 on Premise-Remove old Add2Exchange permissions
                '2' {
                    Clear-Host 
                    'You chose to Remove Permissions to All Users'
                    Write-Host "Removing Old zAdd2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }
                # Option 3: Exchange 2013-2019 on Premise-Remove/Add Permissions all
                '3' {
                    Clear-Host 
                    'You chose to Remove and then Add Permissions to All Users'
                    Write-Host "Removing Old zAdd2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Success....."
                    Write-Host "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Host "Checking............"
                    Write-Host "Done"
                }
                # Option 4: Exchange 2013-2019 on Premise-Adding Permissions to dist. list
                '4' {
                    Clear-Host 
                    'You chose to Add Permissions To A Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Adding Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Add-MailboxPermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                        Write-Host "Done"
                    }
                }
                # Option 5: Exchange 2013-2019 on Premise-Removing dist. list permissions
                '5' {
                    Clear-Host 
                    'You chose to Remove Permissions From A Distribution List'
                    $DistributionGroupName = Read-Host "Enter distribution list name (Display Name)";
                    Write-Host "Removing Add2Exchange Permissions"
                    $DistributionGroupName = Get-DistributionGroupMember -ResultSize Unlimited $DistributionGroupName
                    ForEach ($Member in $DistributionGroupName) {
                        Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                        Write-Host "Done"
                    }
                }
                # Option 6: Exchange 2013-2019 on Premise-Adding permissions to single user
                '6' {
                    Clear-Host 
                    'You chose to Add Permissions To A Single User'
                    $Identity = Read-Host "Enter user Email Address";
                    Write-Host "Adding Add2Exchange Permissions to Single User"
                    Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                    Write-Host "Done" 
                }
                # Option 7: Exchange 2013-2019 on Premise-Removing permissions to single user
                '7' {
                    Clear-Host 
                    'You chose to Remove Permissions To A Single User'
                    $Identity = Read-Host "Enter user Email Address"
                    Write-Host "Removing Add2Exchange Permissions to Single User"
                    Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                    Write-Host "Done"
                }

                # Option 8: Exchange 2013-2019-Adding Permissions to Public Folders
                '8' { 
                    Clear-Host 
                    'You chose to Add Permissions to Public Folders'
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Adding Permissions to Public Folders"
                    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                        Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }
                # Option 9: Exchange 2013-2019-Removing Permissions From Public Folders
                '9' { 
                    Clear-Host
                    'You chose to Remove Permissions From Public Folders' 
                    Write-Host "Getting a list of Public Folders"
                    Get-PublicFolder -Identity "\" -Recurse
                    $Identity = Read-Host "Public Folder Name (Alias)"
                    Write-Host "Removing Permissions to Public Folders"
                    Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
                    Write-Host "Done"
                }
                
                # Option 10: Exchange 2013-2019-Add Permissions to All Public Folders
                '10' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Public Folders'
                    Write-Host "Adding Add2Exchange as Owner to All Public Folders"
                    Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false -ErrorAction SilentlyContinue -ErrorVariable $PError
                    If ($PError) {
                        Get-PublicFolder –Identity "\" –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
                        Get-PublicFolder –Identity "\" –Recurse | Add-PublicFolderClientPermission –User $User –AccessRights Owner -confirm:$false
                    }
                    Write-Host "Done"
                }

                # Option 11: Exchange 2013-2019-Removing Permissions From All Public Folders
                '11' { 
                    Clear-Host 
                    'You chose to Remove Add2Exchange Permissions from All Public Folders'
                    Write-Host "Removing Add2Exchange Owner Permissions from All Public Folders"
                    Get-PublicFolder –Identity “\” –Recurse | Remove-PublicFolderClientPermission –User $User -confirm:$false
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