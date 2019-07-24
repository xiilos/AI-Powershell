if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

<#

Variables:

Office365_Tenent
$TenentUsername
$TenentPassword

Sync_Account
$SyncUsername
$SyncPassword

Distribution_Group_Name
$DistributionGroupName

Dynamic_Distribution_Group
$DynamicDG

Static_Distribution_Group
$StaticDG

Exchange_Server
$ExchangeServerName
$ExchangeUsername
$ExchangePassword

#>



#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Check for Credential Manager Module


$error.clear()
Import-Module "CredentialManager" -ErrorAction SilentlyContinue
If ($error) {
    Write-Host "Adding Credential Manager module"
    Install-Module -Name 'CredentialManager' -WarningAction "Inquire"
} 
Else { Write-Host 'Module is installed' }


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Do {
    
    $Title1 = 'Add2Exchange Task Creation'

    Clear-Host 
    Write-Host "================ $Title1 ================"
    ""
    Write-Host "How Are We Logging In?"
    ""
    Write-Host "Press '1' for Office 365"
    Write-Host "Press '2' for Exchange 2010-2019" 
    Write-Host "Press 'Q' to Quit." -ForegroundColor Red


    #Login Method
 
    $input1 = Read-Host "Please Make A Selection" 
    switch ($input1) { 

        #Office 365--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '1' { 
            Clear-Host 
            'You chose Office 365'
            
            #Create Secure Credentials Office 365 Tenent
                
            $confirmation = Read-Host "Add Office 365 Credentials? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $TenentUsername = Read-Host "Please Type In your Tenent Admin Username"
                $TenentPassword = Read-Host "Please Type in your Tenent Admin Password" -AsSecureString
                $Office365 = @{
                    Target   = "Office365_Tenent"
                    UserName = "$TenentUsername"
                    Password = "$TenentPassword"
                    Comment  = "Office 365 Tenent username and password"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @Office365
            }

            #Create Secure Credentials Sync Service Account
                
            $confirmation = Read-Host "Add Sync Service Account Credentials? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $SyncUsername = Read-Host "Please Type In your Sync Service Account Username"
                $SyncPassword = Read-Host "Please Type in your Sync Service Account Password" -AsSecureString
                $ServiceAccount = @{
                    Target   = "Sync_Account"
                    UserName = "$SyncUsername"
                    Password = "$SyncPassword"
                    Comment  = "Sync Service Account username and password"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @ServiceAccount
            }


            #Create Secure Distribution List Name

            $confirmation = Read-Host "Add Distribution List? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $DistributionGroupName = Read-Host "Please Type In your Office 365 Distribution Group Name"
                    
                $DistributionList = @{
                    Target   = "Distribution_Group_Name"
                    UserName = "$DistributionGroupName"
                    Password = "1234"
                    Comment  = "Office 365 Distribution List Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @DistributionList
            }

            #Create Source Dynamic Distribution List Name

            $confirmation = Read-Host "Add Dynamic Distribution Group? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $DynamicDG = Read-Host "Please Type In your Office 365 Dynamic Distribution Group Name"
                    
                $DynamicDGList = @{
                    Target   = "Dynamic_Distribution_Group"
                    UserName = "$DynamicDG"
                    Password = "1234"
                    Comment  = "Dynamic Distribution Group Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @DynamicDGList
            }

            #Checking Source Static Distribution List Name

            $confirmation = Read-Host "Add Static Distribution Group? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $StaticDG = Read-Host "Please Type In your Office 365 Static Distribution Group Name"
                    
                $StaticDGList = @{
                    Target   = "Static_Distribution_Group"
                    UserName = "$StaticDG"
                    Password = "1234"
                    Comment  = "Static Distribution Group Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @StaticDGList
            }

            #----------------------------------------------------------------------------------------------------------------------------------------------------

            #Check If Tasks Already Exists

            $TaskName = Read-Host "Please Name this Task i.e. Add2Exchange Permissions"

            if (Get-ScheduledTask "$TaskName" -ErrorAction SilentlyContinue) {
                Write-Host "Task Already Exists..."
                $confirmation = Read-Host "Would you like to create a new task? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Write-Host "Creating New Task..."
                }
            }

            #Check for Powershell File Paths

            $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
            Set-Location $Location -ErrorAction SilentlyContinue



            $Title2 = 'Office 365 Task Creation' 
            ""
            Clear-Host 
            Write-Host "================ $Title2 ================" 
            ""     
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Adding Permissions to a Distribution List"
            Write-Host "Press '3' for Adding Permissions to a Dynamic Distribution List" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
    
            
            $input2 = Read-Host "Please Make A Selection" 
            switch ($input2) {

                # Option 1: Office 365-Adding Permissions to All Users
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                
                    $Repeater = (New-TimeSpan -Minutes 360)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to All Users Mailboxes"
                    Write-Host "Done"
                }
                
            
                # Option 2: Office 365-Adding Permissions to a Distribution List
                '2' { 
                    Clear-Host 
                    'You chose to Add Permissions to a Distribution List'
                    $Repeater = (New-TimeSpan -Minutes 360)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dist_List_Permissions.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List"
                    Write-Host "Done" 
                } 

                # Option 3: Office 365-Adding Permissions to a Dynamic Distribution List
                '3' { 
                    Clear-Host 
                    'You chose to Add Permissions to a Dynamic Distribution List'
                    $Repeater = (New-TimeSpan -Minutes 720)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dynamic_Distribution.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Exports a Dynamic Distribution List of Users into a Static List of Users"
                    Write-Host "Done"
                } 
            

                # Option Q: Office 365-Quit
                'q' { 
                    Write-Host "Quitting"
                    Get-PSSession | Remove-PSSession
                    Exit  
                } 
            }

        } 
            
        
        #Exchange2010-2019--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '2' {
            Clear-Host 
            'You chose Exchange 2010-2019'
            
            #Create Secure Credentials Exchange Server
            
            $confirmation = Read-Host "Add Exchange Server Credentials? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
            if ($confirmation -eq 'Y') {
                $ExchangeServerName = Read-Host "Please Type In your Exchange Server Name"
                $ExchangeUsername = Read-Host "Please Type In your Exchange Server Admin Username"
                $ExchangePassword = Read-Host "Please Type in your Exchange Server Admin Password" -AsSecureString
                $ExchangeServer = @{
                    Target   = "Exchange_Server"
                    UserName = "$ExchangeUsername"
                    Password = "$ExchangePassword"
                    Comment  = "$ExchangeServerName"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @ExchangeServer
            }

            #Create Secure Credentials Sync Service Account
                
            $confirmation = Read-Host "Add Sync Service Account Credentials? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $SyncUsername = Read-Host "Please Type In your Sync Service Account Username"
                $SyncPassword = Read-Host "Please Type in your Sync Service Account Password" -AsSecureString
                $ServiceAccount = @{
                    Target   = "Sync_Account"
                    UserName = "$SyncUsername"
                    Password = "$SyncPassword"
                    Comment  = "Sync Service Account username and password"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @ServiceAccount
            }


            #Create Secure Distribution List Name

            $confirmation = Read-Host "Add a Distribution List? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $DistributionGroupName = Read-Host "Please Type In your Distribution Group Name"
                    
                $DistributionList = @{
                    Target   = "Distribution_Group_Name"
                    UserName = "$DistributionGroupName"
                    Password = "1234"
                    Comment  = "Exchange Distribution List Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @DistributionList
            }

            #Create Source Dynamic Distribution List Name

            $confirmation = Read-Host "Add a Dynamic Distribution Group? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $DynamicDG = Read-Host "Please Type In your Exchange Dynamic Distribution Group Name"
                    
                $DynamicDGList = @{
                    Target   = "Dynamic_Distribution_Group"
                    UserName = "$DynamicDG"
                    Password = "1234"
                    Comment  = "Dynamic Distribution Group Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @DynamicDGList
            }

            #Checking Source Static Distribution List Name

            $confirmation = Read-Host "Add a Static Distribution Group? [Y/N]"
            if ($confirmation -eq 'N') {
                Write-Host "Skipping"
            }
   
            if ($confirmation -eq 'Y') {
                $StaticDG = Read-Host "Please Type In your Exchange Static Distribution Group Name"
                    
                $StaticDGList = @{
                    Target   = "Static_Distribution_Group"
                    UserName = "$StaticDG"
                    Password = "1234"
                    Comment  = "Static Distribution Group Name"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @StaticDGList
            }

            #------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            #Check If Tasks Already Exists

            $TaskName = Read-Host "Please Name this Task i.e. Add2Exchange Permissions"

            if (Get-ScheduledTask "$TaskName" -ErrorAction SilentlyContinue) {
                Write-Host "Task Already Exists..."
                $confirmation = Read-Host "Would you like to create a new task? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Write-Host "Creating New Task..."
                }
            }

            #Check for Powershell File Paths

            $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
            Set-Location $Location -ErrorAction SilentlyContinue

            $Title3 = 'Exchange 2010-2019 Task Creation' 
            ""
            Clear-Host 
            Write-Host "================ $Title3 ================" 
            ""     
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Adding Permissions to a Distribution List"
            Write-Host "Press '3' for Adding Permissions to a Dynamic Distribution List" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red

            $input3 = Read-Host "Please Make A Selection" 
            switch ($input3) {

                # Option 1: Exchange 2010-2019-Adding Permissions to All Users
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                    $Repeater = (New-TimeSpan -Minutes 360)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_All_Permissions.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to All Users Mailboxes"
                    Write-Host "Done"
                }

                # Option 2: Exchange 2010-2019-Adding Permissions to a Distribution List
                '2' {
                    Clear-Host 
                    'You chose to Add Permissions to a Distribution List'
                    $Repeater = (New-TimeSpan -Minutes 360)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_Dist_List_Permissions.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List"
                    Write-Host "Done"
                }
                # Option 3: Exchange 2010-2019-Adding Permissions to a Dynamic Distribution List
                '3' {
                    Clear-Host 
                    'You chose to Add Permissions to a Dynamic Distribution List'
                    $Repeater = (New-TimeSpan -Minutes 720)
                    $Duration = ([timeSpan]::maxvalue)
                    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
                    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_Dynamic_Distribution.ps1"'
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Exports a Dynamic Distribution List of Users into a Static List of Users"
                    Write-Host "Done"
                }
            
                # Option Q: Exchange 2010-Quit
                'q' { 
                    Write-Host "Quitting"
                    Get-PSSession | Remove-PSSession
                    Exit 
                }
            }
        }    

    }#Origin LogonMethod End

    $repeat = Read-Host 'Return to the Main Menu? [Y/N]'
} Until ($repeat -eq 'n')

#Quit All
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Start-Sleep -s 1
Exit    
    

# End Scripting