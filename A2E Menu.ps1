if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

##This is a test
function Show-Menu {
    param (
        [string]$Title = 'DidItBetter Software Main Menu'
    )
    Clear-Host
    Write-Host "========================== $Title =========================="
    {

    }
    
    Write-Host "Press '1' Creat Outlook Profile for Add2Exchange."
    Write-Host "NOTE* If you have already installed Outlook on this box then you may proceed. 
     If not; then please install a 32bit version of outlook 2016 for On premise Exchange or Outlook365 if you are on Office 365"
    Write-Host "Note* Make sure that this profile is not in Cache Mode"
    ""
    ""
    Write-Host "Press '2' Give Add2Exchange permissions via powershell for the users you want to sync."
    ""
    ""
    Write-Host "Press '3' Download the latest Add2exchange Software."
    Write-Host "Note* For new installations you will download the full Add2Exchange Enterprise Edition. For Upgrades; you will download the Upgrade versions"
    ""
    ""
    Write-Host "Press '4'Disable User Access Control"
    Write-Host "Note* Reboot after disabling"
    ""
    ""
    Write-Host "Press '5' Setup AutoLogon"
    ""
    ""
    Write-Host "Press '6' Remove Windows 10 Apps"
    ""
    ""
    Write-Host "Press '7' Check Powershell Version"
    ""
    ""
    Write-Host "Press '8' Setup Registry Favorits"
    ""
    ""
    Write-Host "Press '9' Eport License and Profile 1"
    ""
    ""
    Write-Host "Press '10' Remove Outlook Add-Ins"
    ""
    ""
    Write-Host "Press '11' Re-Enable Outlook Add-Ins"
    ""
    ""
    Write-Host "Q: Press 'Q' to quit."
}



do {
    Show-Menu
    ""
    ""
    ""
    $input = Read-Host "Please make a selection"
    switch ($input) {
        '1' {
            Clear-Host
            'You chose option #1'
            Get-ControlPanelItem -Name "Mail (Microsoft Outlook 2016)" | Show-ControlPanelItem
        } '2' {
            Clear-Host
            'You chose option #2'

             
            $message = 'Please Pick how you want to connect'
            $question = 'Pick one of the following from below'

            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange2013-2016'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

            $decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)







            if ($decision -eq 0) {

                Import-Module MSOnline

                Write-Output "Sign in to Office365 as Tenant Admin"
                $Cred = Get-Credential
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic –AllowRedirection
                Import-PSSession $Session
                Import-Module MSOnline
                Connect-MsolService –Credential $Cred


                $message = 'Please Pick what you want to do'
                $question = 'Pick one of the following from below'

                $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Perm O365'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Perm O365'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Add Dist-List Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Remove Dist-List Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&5 Add Single Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&6 Remove Single Perm'))

                $decision = $Host.UI.PromptForChoice($message, $question, $choices, 7)


                if ($decision -eq 0) {

                    $User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

                    Write-Output "Adding Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
                    Invoke-Item "C:\A2E_Office365_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 1) {

                    $User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

                    Write-Output "Removing Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess –verbose -confirm:$false
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
                    Invoke-Item "C:\A2E_Office365_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 2) {

                    $User = read-host "Enter Sync Service Account name Example: zAdd2Exchange or zAdd2Exchange@domain.com";

                    Write-Output "Removing Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess –verbose -confirm:$false
                    Write-Output "Adding Add2Exchange Permissions"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
                    Invoke-Item "C:\A2E_Office365_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 3) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $DistributionGroupName = read-host "Enter distribution list name (Display Name)";

                        Write-Output "Adding Add2Exchange Permissions"

                        $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
                        ForEach ($Member in $DistributionGroupName) {
                            Add-MailboxPermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -AutoMapping:$false
                        }

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 4) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $DistributionGroupName = read-host "Enter distribution list name (Display Name)";

                        Write-Output "Removing Add2Exchange Permissions"

                        $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
                        ForEach ($Member in $DistributionGroupName) {
                            Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -Confirm:$false
                        }

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 5) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $Identity = read-host "Enter user Email Address"

                        Write-Output "Adding Add2Exchange Permissions to Single User"
                        Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                }

                if ($decision -eq 6) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $Identity = read-host "Enter user Email Address"

                        Write-Output "Removing Add2Exchange Permissions to Single User"
                        Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }


            }








            if ($decision -eq 1) {


                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


                Set-ADServerSettings -ViewEntireForest $true


                Write-Output "The next prompt will ask for the Sync Service Account name in the format Example: zAdd2Exchange or zAdd2Exchange@yourdomain.com"
                $User = read-host "Enter Sync Service Account";



                $message = 'Do you Want to remove or Add Add2Exchange Permissions'
                $question = 'Pick one of the following from below'

                $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Both-Remove/Add'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&3 Add Dist-List Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&4 Remove Dist-List Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&5 Add Single Perm'))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&6 Remove Single Perm'))



                $decision = $Host.UI.PromptForChoice($message, $question, $choices, 7)


                if ($decision -eq 0) {
                    Write-Host 'Adding'
                    Write-Output "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Output "Checking............"
                    Start-Sleep -s 2
                    Write-Output "All Done"
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                    Invoke-Item "C:\A2E_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                } 

                if ($decision -eq 1) {
                    Write-Host 'Removing'
                    Write-Output "Removing Old zAdd2Exchange Permissions"
                    Remove-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
                    Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess –verbose -Confirm:$false
                    Write-Output "Checking.............................."
                    Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
                    Write-Output "Success....."
                    Write-Output "Checking............"
                    Start-Sleep -s 2
                    Write-Output "All Done"
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                    Invoke-Item "C:\A2E_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 2) {
                    Write-Host 'Removing'
                    Write-Output "Removing Old zAdd2Exchange Permissions"
                    Remove-ADPermission -Identity “Exchange Administrative Group (FYDIBOHF23SPDLT)” -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
                    Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
                    Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess –verbose -Confirm:$false
                    Write-Output "Checking.............................."
                    Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
                    Write-Output "Success....."
                    Write-Output "Adding Permissions to Users"
                    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
                    Write-Output "Checking............"
                    Start-Sleep -s 2
                    Write-Output "All Done"
                    Write-Output "Writing Data......"
                    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                    Invoke-Item "C:\A2E_permissions.txt"
                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 3) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $DistributionGroupName = read-host "Enter distribution list name (Display Name)";

                        Write-Output "Adding Add2Exchange Permissions"

                        $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
                        ForEach ($Member in $DistributionGroupName) {
                            Add-MailboxPermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -AutoMapping:$false
                        }
                        Write-Output "Writing Data......"
                        Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                        Invoke-Item "C:\A2E_permissions.txt"

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 4) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $DistributionGroupName = read-host "Enter distribution list name (Display Name)";

                        Write-Output "Removing Add2Exchange Permissions"

                        $DistributionGroupName = Get-DistributionGroupMember $DistributionGroupName
                        ForEach ($Member in $DistributionGroupName) {
                            Remove-mailboxpermission -Identity $Member.name -User $User -AccessRights ‘FullAccess’ -InheritanceType all -Confirm:$false
                        }
                        Write-Output "Writing Data......"
                        Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                        Invoke-Item "C:\A2E_permissions.txt"

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 5) {

                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $Identity = read-host "Enter user Email Address"

                        Write-Output "Adding Add2Exchange Permissions to Single User"
                        Add-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false
                        Write-Output "Writing Data......"
                        Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object-Object {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                        Invoke-Item "C:\A2E_permissions.txt"
                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

                if ($decision -eq 6) {
                    do {

                        $User = read-host "Enter Sync Service Account (Display Name)";
                        $Identity = read-host "Enter user Email Address"

                        Write-Output "Removing Add2Exchange Permissions to Single User"
                        Remove-MailboxPermission -Identity $identity -User $User -AccessRights 'FullAccess' -InheritanceType all -Confirm:$false
                        Write-Output "Writing Data......"
                        Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | where {($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | Select-Object Identity, User, @{Name = 'AccessRights'; Expression = {[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
                        Invoke-Item "C:\A2E_permissions.txt"

                        $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                    } Until ($repeat -eq 'n')

                    Write-Output "Quitting"
                    Get-PSSession | Remove-PSSession
                }

            }







        } '3' {
            cls
            'You chose option #3'

            $IE = new-object -com internetexplorer.application
            $IE.navigate2("ftp.diditbetter.com")
            $IE.visible = $true


        } '4' {
            Clear-Host
            'You chose option #4'

            Write-Verbose "Disabling User Access Control"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null
                
        } '5' {
            cls
            'You chose option #5'
            $message = 'Please Pick How to Logon'
            $question = 'Pick one of the following from below'

            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Enable Auto Login'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Disable Auto Login'))

            $decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)



            if ($decision -eq 0) {

                do {
                    $UserName = read-host "Enter the Username (Display Name)";
                    $Password = read-host "Enter the Account password"



                    Write-Verbose "Enabling Windows AutoLogon"
                    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 | out-null
                    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $UserName | out-null
                    New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $Password | out-null

                    $repeat = Read-Host 'Do you want to run it again? [Y/N]'

                } Until ($repeat -eq 'n')

                Write-Output "Quitting"
                Get-PSSession | Remove-PSSession
            }




            if ($decision -eq 1) {


                Write-Verbose "Disabling Windows AutoLogon"
                Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon | out-null
                Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName | out-null
                Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword | out-null

                Write-Output "Quitting"
                Get-PSSession | Remove-PSSession
            }


        } '6' {
            cls
            'You chose option #6'
            Write-Verbose "Removing Windows 10 Apps"
            Get-AppxPackage | Where-Object-object {$_.name –notlike "*photos"} | Where-Object-object {$_.name –notlike "*store*"} | Where-Object-Object-Object-Object-object {$_.name –notlike "*windowscalculator*"} | Remove-AppxPackage -Confirm:$False
                
        } '7' {
            cls
            'You chose option #7'
            Write-Verbose "Checking Powershell Version"
            Get-Host
                
        } '8' {
            cls
            'You chose option #8'
            Write-Verbose "Adding Registry Favorits"
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name Session Manager -Value Computer\\HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name EnableLUA -Value Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name .NetFramework -Value Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\.NETFramework
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name OpenDoor Software -Value Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\OpenDoor Software®
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name PendingFileRename -Value Computer\\HKEY_LOCAL_MACHINE\\SYSTEM\\ControlSet001\\Control\\Session Manager
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name AutoDiscover -Value HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\16.0\\Outlook\\AutoDiscover
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name Office -Value Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System
            Get-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" | New-ItemProperty -Name EnableLUA -Value Computer\\HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem

                
        } '9' {
            Clear-Host
            'You chose option #9'
            write-verbose "Exporting License and Profile 1"
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" C:\zlibrary\License_Info.Reg
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" C:\zlibrary\Profile_1.Reg


        } '10' {
            Clear-Host
            'You chose option #10'
            write-verbose "Remove Outlook Add-Ins"
            Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALCONNECTOR.DLL*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALCONNECTOR.DLL', 'SOCIALCONNECTORbckup.dll' }
            Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16" "*SOCIALPROVIDER.DLL*" -Recurse | Rename-Item -NewName {$_.name -replace 'SOCIALPROVIDER.DLL', 'SOCIALPROVIDERbckup.dll' }
            Get-ChildItem -path "C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS" "*ColleagueImport.dll*" -Recurse | Rename-Item -NewName {$_.name -replace 'ColleagueImport.dll', 'ColleagueImportbckup.dll' }

        } '11' {
            Clear-Host
            'You chose option #11'
            write-verbose "Re-Enable Outlook Add-Ins"

           
        } 'q' {
            return
        }
    }
    pause
}
until ($input -eq 'q')
