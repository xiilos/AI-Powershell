if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #
Do {

    $Title1 = 'Add2Exchange Post Migration Wizard'

    Clear-Host 
    Write-Host "================ $Title1 ================"
    ""
    Write-Host "Please Pick What Best Fits your Scenario"
    ""
    Write-Host "Press '1' for Add2Exchange is Not Installed on This Box yet."
    Write-Host "Press '2' for Add2Exchange is Already Installed and the Console has already been opened Once" 
    Write-Host "Press 'Q' to Quit." -ForegroundColor Red 

    #Migration Method
 
    $input1 = Read-Host "Please Make A Selection" 
    switch ($input1) { 

        #POST Automatic Migration--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '1' { 
            Clear-Host 
            'You chose The Add2Exchange is Not Installed on this Box Yet Option'
        
            #Check for UAC First

            # Disable UAC
            Write-Host "Checking UAC"

            $Val = Get-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA"

            if ($val.EnableLUA -ne 0) {
                Set-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA" -value 0
                Write-Host "Done"
                Write-Host "UAC is now Disabled. Please Reboot and Run this Script Again" -ForegroundColor Red
                Pause
                Get-PSSession | Remove-PSSession
                Exit

            }

            Else {

                Write-Host "UAC is already Disabled"
                Write-Host "Resuming"
            }


            #Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1
            # Account Creation

            $message = 'Have you created an account for Add2Exchange to install under?'
            $question = 'Pick one of the following from below'

            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

            $decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

            # Option 2: Quit

            if ($decision -eq 2) {
                Write-Host "Quitting"
                Get-PSSession | Remove-PSSession
                Exit
            }


            if ($decision -eq 0) {


                # Account Creation

                Write-host "We need to create an account for Add2Exchange"
                $confirmation = Read-Host "Are you on a domain? [Y/N]"
                if ($confirmation -eq 'y') {

                    $wshell = New-Object -ComObject Wscript.Shell
                    $answer = $wshell.Popup("Please create a new account called zadd2exchange (or any name of your choosing).
Note* The account only needs Domain User and Public Folder management permissions.
Note* You cannot hide this account.
Once Done, add the new user account as a local Administrator of this box.
Log off and back on as the new Sync Service account and run this again before you Proceed.", 0, "Account Creation", 0x1)
                    if ($answer -eq 2) { Break }

                    $confirmation = Read-host "Do you want to continue [Y/N]"
                    if ($confirmation -eq 'y') {
                        Write-Host "Resumming"
                    }

                    if ($confirmation -eq 'n') {
                        Write-Host "Quitting"
                        Get-PSSession | Remove-PSSession
                        Exit
                    }
                }



                if ($confirmation -eq 'n') {
                    Write-Host "Lets create a new user for Add2Exchange"
                    $Username = Read-Host "Username for new account"
                    $Fullname = Read-host "User Full Name"    
                    $password = Read-Host "Password for new account" -AsSecureString
                    $description = read-host "Description of this account."

                    New-LocalUser "$Username" -Password $Password -FullName "$Fullname" -Description "$description"
                    Add-LocalGroupMember -Group "Administrators" -Member $Username

                }
            }

            if ($decision -eq 1) {
                $confirmation = Read-Host "Are you logged into this box as the new sync account? [Y/N]"
                if ($confirmation -eq 'y') {
                    Write-Host "Resumming"
                }

    
                if ($confirmation -eq 'n') {  
                    $wshell = New-Object -ComObject Wscript.Shell
                    $answer = $wshell.Popup("Add the new user account as a local Administrator of this box.
Log off and back on as the new Sync Service account and run this again before proceeding.", 0, "Account Creation", 0x1)
                    if ($answer -eq 2) { Break }

                    $confirmation = Read-host "Do you want to continue [Y/N]"
                    if ($confirmation -eq 'y') {
                        Write-Host "Resumming"
                    }

                    if ($confirmation -eq 'n') {
                        Write-Host "Quitting"
                        Get-PSSession | Remove-PSSession
                        Exit
                    }

                }
            }

            # Step 2-----------------------------------------------------------------------------------------------------------------------------------------------------Step 2

            # Powershell Update
            # Check if .Net 4.5 or above is installed

            $release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
            $installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

            if (($installed -ne 1) -or ($release -lt 378389)) {
                Write-Host "We need to download .Net 4.5.2"
                Write-Host "Downloading"
                $Directory = "C:\PowerShell"

                if ( -Not (Test-Path $Directory.trim() )) {
                    New-Item -ItemType directory -Path C:\PowerShell
                }

                $url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                $output = "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                (New-Object System.Net.WebClient).DownloadFile($url, $output)    
                Start-Process -FilePath "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -wait
                Write-Host "Download Complete"
                $wshell = New-Object -ComObject Wscript.Shell
                $answer = $wshell.Popup("Please Reboot after Installing and run this again", 0, "Done", 0x1)
                if ($answer -eq 2) { Break }
                Write-Host "Quitting"
                Get-PSSession | Remove-PSSession
                Exit
            }


            #Check Operating Sysetm

            $BuildVersion = [System.Environment]::OSVersion.Version


            #OS is 10+
            if ($BuildVersion.Major -like '10') {
                Write-Host "WMF 5.1 is not supported for Windows 10 and above"
            }

            #OS is 7
            if ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1') {
        
                Write-Host "Downloading WMF 5.1 for 7+"
                $Directory = "C:\PowerShell"

                if ( -Not (Test-Path $Directory.trim() )) {
                    New-Item -ItemType directory -Path C:\PowerShell
                }
                $url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
                $output = "C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu"
                (New-Object System.Net.WebClient).DownloadFile($url, $output)
                Start-Process -FilePath 'C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu' -wait
                Write-Host "Download Complete"
                $wshell = New-Object -ComObject Wscript.Shell
                $answer = $wshell.Popup("Please Reboot after Installing and run this again", 0, "Done", 0x1)
                if ($answer -eq 2) { Break }
                Write-Host "Quitting"
                Get-PSSession | Remove-PSSession
                Exit
            }

            #OS is 8
            elseif ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3') {
        
                Write-Host "Downloading WMF 5.1 for 8+"
                $Directory = "C:\PowerShell"

                if ( -Not (Test-Path $Directory.trim() )) {
                    New-Item -ItemType directory -Path C:\PowerShell
                }
                $url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
                $output = "C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu"
                (New-Object System.Net.WebClient).DownloadFile($url, $output)
                Start-Process -FilePath 'C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu' -wait
                Write-Host "Download Complete"
                $wshell = New-Object -ComObject Wscript.Shell
                $answer = $wshell.Popup("Please Reboot after Installing and run this again", 0, "Done", 0x1)
                if ($answer -eq 2) { Break }
                Write-Host "Quitting"
                Get-PSSession | Remove-PSSession
                Exit
            }

            Write-Host "Nothing to do"
            Write-Host "You Are on the latest version of PowerShell"


            #Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3

            #Create zLibrary & Copy Shortcuts to Desktop

            $TestPath = "C:\zlibrary"

            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

                Write-Host "zLibrary exists...Resuming"
            }
            Else {
                New-Item -ItemType directory -Path C:\zlibrary
            }

            # Desktop Shortcuts

            $wshShell = New-Object -ComObject "WScript.Shell"
            $urlShortcut = $wshShell.CreateShortcut(
                (Join-Path $wshShell.SpecialFolders.Item("AllUsersDesktop") "Disable Outlook Social Connector through GPO.url")
            )
            $urlShortcut.TargetPath = "https://support.microsoft.com/en-us/help/2020103/how-to-manage-the-outlook-social-connector-by-using-group-policy"
            $urlShortcut.Save()


            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\zLibrary.lnk")
            $Shortcut.TargetPath = "C:\zlibrary"
            $Shortcut.Save()


            # Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4

            # Outlook Install 32Bit

            If (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application -ErrorAction SilentlyContinue) {
                Write-Output "Outlook Is Installed"
            }
            Else {
                Write-host "We need to install Outlook on this machine"
                $confirmation = Read-Host "Would you like me to Install Outlook 365? [Y/N]"
                if ($confirmation -eq 'y') {
                    #Downloading Office 365

                    Write-Host "Downloading Office 365 Files"
                    Write-Host "Please Wait......"

                    $URL = "ftp://ftp.diditbetter.com/Outlook%20365/O365Outlook32.zip"
                    $Output = "c:\zlibrary\O365Outlook32.zip"
                    $Start_Time = Get-Date

                    (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

                    Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

                    Write-Host "Finished Downloading"

                    #Unpacking Office 365

                    Write-Host "Unpacking O365 Files"
                    Write-Host "please Wait....."
                    Push-Location "c:\zlibrary\"
                    Expand-Archive -Path "c:\zlibrary\O365Outlook32.zip" -DestinationPath "c:\zlibrary\Office 365"
                    Start-Sleep -Seconds 2
                    Write-Host "Installing Office 365. Please Wait..."
                    Push-Location -Path "c:\zlibrary\Office 365"
                    .\setup.exe /configure Office365x32bit.xml
                    Write-Host "Done"

                }

                if ($confirmation -eq 'n') {

                    $wshell = New-Object -ComObject Wscript.Shell
                    $answer = $wshell.Popup("If you choose to install your own Outlook; ensure that it is Outlook 2016 and above, and 32bit Only.
When Done, click OK", 0, "Outlook Install", 0x1)
                    if ($answer -eq 2) { Break }

                }
            }


            # Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5

            # Mailbox Creation

            $confirmation = Read-Host "Are you on Office 365 or Exchange on Premise [O/E]"
            if ($confirmation -eq 'O') {
                Write-host "Make sure to create a mailbox for the sync service account and add an *E* License to it"
                $wshell = New-Object -ComObject Wscript.Shell

                $answer = $wshell.Popup("When this is Done, click OK to Continue", 0, "Create a Mailbox", 0x1)
                if ($answer -eq 2) { Break }
            }

            if ($confirmation -eq 'E') {
                Write-Host "Create a Mailbox for the new Sync Service account in Exchange"
                $wshell = New-Object -ComObject Wscript.Shell

                $answer = $wshell.Popup("When this is Done, Click OK to Continue", 0, "Create a Mailbox", 0x1)
                if ($answer -eq 2) { Break }
            }

            # Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6

            # Mail Profile

            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("The next step is to Create a Profile for your new account. Open Control panel and go to Mail. Create a new profile and follow through the steps that pertain to your Organization. 
Note* Make sure you do NOT have Cache checked. When this is finished click OK to Continue", 0, "Creating an Outlook Profile", 0x1)
            if ($answer -eq 2) { Break }


            # Step 7-----------------------------------------------------------------------------------------------------------------------------------------------------Step 7
            
            # Checking for Install files and Download if needed

            Write-Host "Pick the A2E_Backup Folder from the Browser"
            $Application = New-Object -ComObject Shell.Application
            $Path = ($application.BrowseForFolder(0, 'Select a folder', 0)).Self.Path
            Write-Host "You Chose $Path"
            $TestPath = "$Path\a2e-enterprise.exe"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

                Write-Host "Add2Exchange Install Files Exists...Resuming"
                Move-Item -Path "$Path\a2e-enterprise.exe" -Destination "c:\zlibrary"

                #Unpacking Add2Exchange
                Write-Host "Unpacking Add2exchange"
                Write-Host "please Wait....."
                Push-Location "c:\zlibrary"
                Start-Process "c:\zlibrary\a2e-enterprise.exe" -wait
                Start-Sleep -Seconds 2
                Write-Host "Done"

            }
            Else {
                Write-Host "Downloading Add2Exchange Enterprise"
                Write-Host "Please Wait......"

                $URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/New%20Installs/a2e-enterprise.exe"
                $Output = "C:\zlibrary\A2E-Enterprise.exe"
                $Start_Time = Get-Date

                (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

                Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

                Write-Host "Finished Downloading"

                #Unpacking Add2Exchange

                Write-Host "Unpacking Add2exchange"
                Write-Host "please Wait....."
                Push-Location "c:\zlibrary"
                Start-Process "c:\zlibrary\a2e-enterprise.exe" -wait
                Start-Sleep -Seconds 2
                Write-Host "Done"

            }


            # Step 8-----------------------------------------------------------------------------------------------------------------------------------------------------Step 8
            #Installing Add2Exchange
            Write-Host "Installing Add2Exchange"
            $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Push-Location $Location
            Do {
                Start-Process -FilePath ".\Add2ExchangeSetup.msi" -wait -ErrorAction SilentlyContinue -ErrorVariable InstallError;
            
                If ($InstallError) { 
                    Write-Warning -Message "Something Went Wrong with the Install!"
                    Write-Host "Trying The Install Again in 2 Seconds"
                    Start-Sleep -S 2
                }
            } Until (-not($InstallError))
            Write-Host "Done"

            # Step 9-----------------------------------------------------------------------------------------------------------------------------------------------------Step 9
            Write-Host "Installing SQL"
            Start-Process "$Home\Desktop\Add2exchange Console"
           
            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Once the Install is complete, Click OK to finish the setup", 0, "Finishing Installation", 0x1)
            if ($answer -eq 2) { Break }

            Stop-Process -Name "Add2Exchange Console"

            # Step 10-----------------------------------------------------------------------------------------------------------------------------------------------------Step 10
            # Registry Favorites & Shortcuts

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


            # Step 11-----------------------------------------------------------------------------------------------------------------------------------------------------Step 11
            # Completion and Wrap-Up

            #Removing Outlook Social Connector
            Write-Host "Removing Outlook Social Connector"
            Start-Process ".\Setup\OSC_Disable.bat"


            # Step 12-----------------------------------------------------------------------------------------------------------------------------------------------------Step 12
            Write-Host "Migrating Add2Exchange SQL Database. Please Wait....."
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
           
            Rename-Item -Path "$Install\Database\A2E.mdf" -NewName "Clean_A2E.mdf" -ErrorAction SilentlyContinue
            Copy-Item "$Path\A2E.mdf" -Destination "$Install\Database" -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error Copying File A2E.mdf. Please Copy the File Manually into $Install\Database"
                Pause
            }
            Rename-Item -Path "$Install\Database\A2E_log.LDF" -NewName "Clean_A2E_log.LDF" -ErrorAction SilentlyContinue
            Copy-Item "$Path\A2E_log.LDF" -Destination "$Install\Database" -ErrorAction SilentlyContinue -ErrorVariable DB2
            If ($DB2) {
                Write-Host "Error Copying File A2E_Log.LDF. Please Copy the File Manually into $Install\Database"
                Pause
            }
            Start-Service -Name "SQL Server (A2ESQLSERVER)"
            Write-Host "Done"

            # Step 13-----------------------------------------------------------------------------------------------------------------------------------------------------Step 13
            Write-Host "Getting Required Files... Please Wait.."
            
            Copy-Item "$Path\Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            Copy-Item "$Path\Old_Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            Copy-Item "$Path\Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware" -ErrorAction SilentlyContinue
            Copy-Item "Path\DidItBetterSoftware\*" -Destination "C:\Program Files (x86)" -Recurse -ErrorAction SilentlyContinue
            REG Import $Path\License_Info.Reg
            

            #Adding Permissions
            $Confirmation = Read-Host "Do We need to run through permissions? [Y/N]"
            if ($confirmation -eq 'y') {
                Push-Location $Install
                Start-Process Powershell .\Setup\PermissionsOnPremOrO365Combined.ps1
            }
            if ($confirmation -eq 'n') {
                Write-Host "Skipping"
            }

            #Adding AutoLogon
            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue", 0, "AutoLogin", 0x1)
            if ($answer -eq 2) { Break }

            Start-Process -FilePath ".\Setup\AutoLogon.exe" -wait

            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Setup is Complete. You can now start the Add2Echange Console", 0, "Done", 0x1)
            if ($answer -eq 2) { Break }

        }
            

        '2' { 
            Clear-Host 
            'You chose Add2Exchange is Already Installed and the Console has already been opened Once'

            Stop-Service -Name "SQL Server (A2ESQLSERVER)" -ErrorAction SilentlyContinue

            # Registry Favorites & Shortcuts

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


            
            # Completion and Wrap-Up

            #Removing Outlook Social Connector
            Write-Host "Removing Outlook Social Connector"
            Start-Process ".\Setup\OSC_Disable.bat"


            #Migrating A2E SQL DB
            Write-Host "Migrating Add2Exchange SQL Database. Please Wait....."
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
           
            Rename-Item -Path "$Install\Database\A2E.mdf" -NewName "Clean_A2E.mdf" -ErrorAction SilentlyContinue
            Copy-Item "$Path\A2E.mdf" -Destination "$Install\Database" -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error Copying File A2E.mdf. Please Copy the File Manually into $Install\Database"
                Pause
            }
            Rename-Item -Path "$Install\Database\A2E_log.LDF" -NewName "Clean_A2E_log.LDF" -ErrorAction SilentlyContinue
            Copy-Item "$Path\A2E_log.LDF" -Destination "$Install\Database" -ErrorAction SilentlyContinue -ErrorVariable DB2
            If ($DB2) {
                Write-Host "Error Copying File A2E_Log.LDF. Please Copy the File Manually into $Install\Database"
                Pause
            }
            Start-Service -Name "SQL Server (A2ESQLSERVER)"
            Write-Host "Done"

            #Moving Files
            Write-Host "Getting Required Files... Please Wait.."
            
            Copy-Item "$Path\Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            Copy-Item "$Path\Old_Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            Copy-Item "$Path\Support.txt" -Destination "C:\zlibrary" -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware" -ErrorAction SilentlyContinue
            Copy-Item "Path\DidItBetterSoftware\*" -Destination "C:\Program Files (x86)" -Recurse -ErrorAction SilentlyContinue
            REG Import $Path\License_Info.Reg
            

            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Setup is Complete. You can now start the Add2Echange Console", 0, "Done", 0x1)
            if ($answer -eq 2) { Break }

            #Adding Permissions
            $Confirmation = Read-Host "Do We need to run through permissions? [Y/N]"
            if ($confirmation -eq 'y') {
                Push-Location $Install
                Start-Process Powershell .\Setup\PermissionsOnPremOrO365Combined.ps1
            }
            if ($confirmation -eq 'n') {
                Write-Host "Skipping"
            }

            #Adding AutoLogon
            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue", 0, "AutoLogin", 0x1)
            if ($answer -eq 2) { Break }

            Start-Process -FilePath ".\Setup\AutoLogon.exe" -wait

        }

        # Option Q: Migration Wizard - Quit
        'q' { 
            Write-Host "Quitting"
            Get-PSSession | Remove-PSSession
            Exit
        }

    }
    Write-Host "Done"
    $repeat = Read-Host 'Return To The Main Menu? [Y/N]'
} Until ($repeat -eq 'n')


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting