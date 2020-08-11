if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Goal
# Create Initial Environment for Add2Exchange Install
# Assign Permissions for Add2Exchange
# Install Add2Exchange
# Cleanup
# Luanch Add2Exchange for the first time


#Step 1: Account Creation
#Step 2: Upgrade .Net and Powershell if needed
#Step 3: Create zLibrary and Create Shortcuts
#Step 4: Install Outlook and Setup Profile
#Step 5: Mailbox Creation
#Step 6: Create a Mail Profile
#Step 7: Add Permissions
#Step 8: Add Public Folder Permissions
#Step 9: Enable AutoLogon
#Step 10: Install Add2Exchange
#Step 11: Add Registry Favs
#Step 12: Cleanup

#Check for UAC First

# Disable UAC
$Title1 = 'Add2Exchange Installation Wizard'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "Please Pick from the below"
""
Write-Host "Press '1' to Start the Installation Wizard"
Write-Host "Press 'Q' to Quit." -ForegroundColor Red 

#Migration Method
 
$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    # Start of Automated Scripting #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    '1' { 
        Clear-Host 
        'You chose The Add2Exchange Installation Wizard'

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
        # Check Operating Sysetm

        $Confirmation = Read-Host "Check to see if you are on the latest powershell? [Y/N]"
        if ($confirmation -eq 'y') {
        
            Start-Process Powershell ".\Setup\Legacy_PowerShell.ps1" -wait

        }
        if ($confirmation -eq 'n') {
            Write-Host "Skipping"
        }


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

                Write-Host "Please Wait while we install Office 365 Pro Retail"
                
                Push-Location -Path ".\O365Outlook32\Setup Files"
                .\setup.exe /configure Office365_Pro_Reatilx86_Configuration.xml
                Pop-Location
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
Note* Make sure you do not have Cache checked. When this is finished click OK to Continue", 0, "Creating an Outlook Profile", 0x1)
        if ($answer -eq 2) { Break }

        # Step 7-----------------------------------------------------------------------------------------------------------------------------------------------------Step 7

        #Adding Permissions
        $Confirmation = Read-Host "Do We need to run through Add2Exchange permissions? [Y/N]"
        if ($confirmation -eq 'y') {
        
            Start-Process Powershell ".\Setup\PermissionsOnPremOrO365Combined.ps1" -wait
            
        }
        if ($confirmation -eq 'n') {
            Write-Host "Skipping"
        }

        # Step 8-----------------------------------------------------------------------------------------------------------------------------------------------------Step 8
        # Auto Logon

        $wshell = New-Object -ComObject Wscript.Shell

        $answer = $wshell.Popup("Now we can set the AutoLogon feature for this account.
Note* Please fill in all areas on the next screen to enable Auto logging on to this box.
Click OK to Continue or Cancel to Skip", 0, "AutoLogin", 0x1)

        if ($answer -eq 1) {
            Start-Process -FilePath ".\Setup\AutoLogon.exe" -wait
        }         

        if ($answer -eq 2) { 
            Write-Host "Skipping AutoLogon"
        }


        # Step 9-----------------------------------------------------------------------------------------------------------------------------------------------------Step 9
        # Installing the Software

        $wshell = New-Object -ComObject Wscript.Shell

        $answer = $wshell.Popup("System Setup Complete. Lets Install the Software", 0, "Complete", 0x1)
        if ($answer -eq 2) { Break }
        Do {
        
            Start-Process -FilePath ".\Add2ExchangeSetup.msi" -wait -ErrorAction SilentlyContinue -ErrorVariable InstallError;

            If ($InstallError) { 
                Write-Warning -Message "Something Went Wrong with the Install!"
                Write-Host "Trying The Install Again in 2 Seconds"
                Start-Sleep -S 2
            }
        } Until (-not($InstallError))


        Start-Process "$Home\Desktop\Add2exchange Console"
        $wshell = New-Object -ComObject Wscript.Shell

        $answer = $wshell.Popup("Once the Install is complete, Click OK to finish the setup", 0, "Finishing Installation", 0x1)
        if ($answer -eq 2) { Break }


        # Step 10-----------------------------------------------------------------------------------------------------------------------------------------------------Step 11
        # Registry Favorites & Shortcuts

        Start-Process Powershell ".\Setup\Registry_Favorites.ps1" -wait

        # Step 11-----------------------------------------------------------------------------------------------------------------------------------------------------Step 11
        # Completion and Wrap-Up

        #Removing Outlook Social Connector
        Write-Host "Removing Outlook Social Connector"
    
        Start-Process Powershell ".\Setup\OSC_Disable.bat" -wait

        #Creating Setup Details
    
        Start-Process Powershell ".\Setup\A2E_Setup_Details.ps1" -wait


        $wshell = New-Object -ComObject Wscript.Shell

        $answer = $wshell.Popup("Setup is Complete. You can now use Add2Exchange", 0, "Done", 0x1)
        if ($answer -eq 2) { Break }
        Write-Host "ttyl"
        Get-PSSession | Remove-PSSession
        Exit
    }
    
        
    # Option Q: Migration Wizard - Quit
    'q' { 
        Write-Host "Quitting"
        Get-PSSession | Remove-PSSession
        Exit
    }
    
}
# End Scripting