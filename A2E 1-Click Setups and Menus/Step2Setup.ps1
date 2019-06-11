if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Goal
# Assign Permissions for Add2Exchange
# Install Add2Exchange
# Cleanup
# Luanch Add2Exchange for the first time

#Step 1: Add Permissions
#Strp 2: Add Public Folder Permissions
#Step 3: Enable AutoLogon
#Step 4: Install Add2Exchange
#Step 5: Add Registry Favs
#Step 6: Cleanup

# Start of Automated Scripting #

# Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1

# Office 365 and on premise Exchange Permissions

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("In this step, we will assign the service account full mailbox access to the users that will be syncing with Add2Exchange. 
Instead of adding permissions to everyone, create a Distribution list with all users that will be syncing with Add2Exchange. 
Reminder*** You cannot hide this Distribution list, so it helps to put a Z in front of it to drop it to the bottom of the GAL", 0, "Assigning Mailbox Permissions", 0x1)
if ($answer -eq 2) { Break }

$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}

Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_Permissions.txt" -Append

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
            If (!$Cred) { Exit }
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
                    $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
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
  

# Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3
# Auto Logon

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Excellent!! Permissions are done and now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue", 0, "AutoLogin", 0x1)
if ($answer -eq 2) { Break }

Start-Process -FilePath ".\Setup\AutoLogon.exe" -wait


# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4
# Installing the Software

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("System Setup Complete. Lets Install the Software", 0, "Complete", 0x1)
if ($answer -eq 2) { Break }
Do {
    Start-Process -FilePath ".\Add2ExchangeSetup.msi" -wait -ErrorAction Inquire -ErrorVariable InstallError;

    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))



# Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5
# Registry Favorites & Shortcuts

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Once the Install is complete, Click OK to finish the setup", 0, "Finishing Installation", 0x1)
if ($answer -eq 2) { Break }

Write-Host "Creating Registry Favorites"
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Session Manager" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControl\Control\Session Manager" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "EnableLUA" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name ".Net Framework" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "OpenDoor Software" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Add2Exchange"  -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Pendingfilerename" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "AutoDiscover" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Office" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Group Policy History" -Type string -Value "Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\History" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Logon" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Update" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Outlook Social Connector" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "MFA1" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Exchange" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "MFA2" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Force -ErrorAction SilentlyContinue


# Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6
# Completion and Wrap-Up

#Removing Outlook Social Connector

$TestPath = "C:\Program Files (x86)\Microsoft Office\root\Office16"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    Write-Host "32bit Outlook"
    Write-Host "Setting Load Behavior to 0"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

}
Else {
    Write-Host "64bit Outlook"
    Write-Host "Setting Load Behavior to 0"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

}

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("Setup is Complete. You can now start the Add2Echange Console", 0, "Done", 0x1)
if ($answer -eq 2) { Break }
Get-PSSession | Remove-PSSession
Exit
# End Scripting