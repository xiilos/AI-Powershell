if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Check and Create Stored Credentials
Do {

    $TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
    if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
        Write-Host "Secure Location Exists...Resuming"
    }
    Else {
        Write-Host "Creating Secure Location"
        New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
    }

    Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
            #Create Secure Credentials File

            Write-Host "We will now create your Secure Credential Files"
            #Checking Source Tenent or Exchange Admin Username

            $TestPath = ".\ServerUser.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange/Tenent Admin Username File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerUser.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Global Admin Username or Exchange Admin Username (Used to Log Into Office 365 or On Premise Exchange)" | out-file ".\ServerUser.txt"
                }

            }

            Else {
                Read-Host "Global Admin Username or Exchange Admin Username (Used to Log Into Office 365 or On Premise Exchange)" | out-file ".\ServerUser.txt"
            }
   
            #Checking Source Tenent or Exchange Admin Password

            $TestPath = ".\ServerPass.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange/Tenent Admin Password File Exists..."
                Write-Host "Last Updated:" -ForegroundColor Green
                Get-Item "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerPass.txt" | ForEach-Object { $_.LastWriteTime }
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Global Admin Password or Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
                }

            }

            Else {
                Read-Host "Global Admin Password or Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
            }
   

            #Checking Source Exchange Name

            $TestPath = ".\ExchangeName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange Server Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchangename.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
       
                if ($confirmation -eq 'Y') {
                    Read-Host "If on Premise; Type in your Exchange Server Name. Leave Blank for None. Press Enter when Finished." | out-file ".\ExchangeName.txt"
                }
            }
   
            Else {
                Read-Host "If on Premise; Type in your Exchange Server Name. Leave Blank for None. Press Enter when Finished." | out-file ".\ExchangeName.txt"
            }

            #Checking Source Distribution List Name

            $TestPath = ".\DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "If Adding Permissions to a Distribution List Type in the Distribution List Name. Leave Blank for None. Press Enter when Finished." | out-file ".\DistributionName.txt"
                }
            }

            Else {
                Read-Host "If Adding Permissions to a Distribution List Type in the Distribution List Name. Leave Blank for None. Press Enter when Finished." | out-file ".\DistributionName.txt"
            }

            #Checking Source Dynamic Distribution List Name

            $TestPath = ".\Dynamic_DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Dynamic Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dynamic_DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Dynamic Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
                Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
                Leave Blank for None. Press Enter when Finished." | out-file ".\Dynamic_DistributionName.txt"
                }
            }

            Else {
                Read-Host "Type in the Dynamic Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
                Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
                Leave Blank for None. Press Enter when Finished." | out-file ".\Dynamic_DistributionName.txt"
            }

            #Checking Source Static Distribution List Name

            $TestPath = ".\Static_DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Statis Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Static_DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Static Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
                Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
                Leave Blank for None. Press Enter when Finished." | out-file ".\Static_DistributionName.txt"
                }
            }

            Else {
                Read-Host "Type in the Static Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
                Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
                Leave Blank for None. Press Enter when Finished." | out-file ".\Static_DistributionName.txt"
            }

            #Checking Source Service Account Name

            $TestPath = ".\ServiceAccount.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Service Account Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServiceAccount.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Sync Service Account Name. Example: *zAdd2Exchange* Press Enter when Finished." | out-file ".\ServiceAccount.txt"
                }
            }

            Else {
                Read-Host "Type in the Sync Service Account Name. Example: *zAdd2Exchange* Press Enter when Finished." | out-file ".\ServiceAccount.txt"
            }


            #Check If Tasks Already Exists

            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists..."
                $confirmation = Read-Host "Would You Like to Update the Current Task? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
                }
            }

            #Check for Powershell File Paths

            $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
            Set-Location $Location



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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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
            #Create Secure Credentials File

            Write-Host "We will now create your Secure Credential Files"
            #Checking Source Tenent or Exchange Admin Username

            $TestPath = ".\ServerUser.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange/Tenent Admin Username File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerUser.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Global Admin Username or Exchange Admin Username (Used to Log Into Office 365 or On Premise Exchange)" | out-file ".\ServerUser.txt"
                }

            }

            Else {
                Read-Host "Global Admin Username or Exchange Admin Username (Used to Log Into Office 365 or On Premise Exchange)" | out-file ".\ServerUser.txt"
            }
   
            #Checking Source Tenent or Exchange Admin Password

            $TestPath = ".\ServerPass.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange/Tenent Admin Password File Exists..."
                Write-Host "Last Updated:" -ForegroundColor Green
                Get-Item "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerPass.txt" | ForEach-Object { $_.LastWriteTime }
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Global Admin Password or Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
                }

            }

            Else {
                Read-Host "Global Admin Password or Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
            }
   

            #Checking Source Exchange Name

            $TestPath = ".\ExchangeName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Exchange Server Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchangename.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
       
                if ($confirmation -eq 'Y') {
                    Read-Host "If on Premise; Type in your Exchange Server Name. Leave Blank for None. Press Enter when Finished." | out-file ".\ExchangeName.txt"
                }
            }
   
            Else {
                Read-Host "If on Premise; Type in your Exchange Server Name. Leave Blank for None. Press Enter when Finished." | out-file ".\ExchangeName.txt"
            }

            #Checking Source Distribution List Name

            $TestPath = ".\DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "If Adding Permissions to a Distribution List Type in the Distribution List Name. Leave Blank for None. Press Enter when Finished." | out-file ".\DistributionName.txt"
                }
            }

            Else {
                Read-Host "If Adding Permissions to a Distribution List Type in the Distribution List Name. Leave Blank for None. Press Enter when Finished." | out-file ".\DistributionName.txt"
            }

            #Checking Source Dynamic Distribution List Name

            $TestPath = ".\Dynamic_DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Dynamic Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dynamic_DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Dynamic Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
        Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
        Leave Blank for None. Press Enter when Finished." | out-file ".\Dynamic_DistributionName.txt"
                }
            }

            Else {
                Read-Host "Type in the Dynamic Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
        Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
        Leave Blank for None. Press Enter when Finished." | out-file ".\Dynamic_DistributionName.txt"
            }

            #Checking Source Static Distribution List Name

            $TestPath = ".\Static_DistributionName.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Statis Distribution List Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Static_DistributionName.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Static Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
        Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
        Leave Blank for None. Press Enter when Finished." | out-file ".\Static_DistributionName.txt"
                }
            }

            Else {
                Read-Host "Type in the Static Distribution List Name. For Multiple Entries Example: zFirmContacts,zFirmCalendar
        Note: Make sure to use qoutes  arounf each groupseperated by a comma. 
        Leave Blank for None. Press Enter when Finished." | out-file ".\Static_DistributionName.txt"
            }

            #Checking Source Service Account Name

            $TestPath = ".\ServiceAccount.txt"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
                Write-Host "Service Account Name File Exists..."
                Write-Host "Current Content of File:" -ForegroundColor Green
                Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServiceAccount.txt"
                ""
                $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }
   
                if ($confirmation -eq 'Y') {
                    Read-Host "Type in the Sync Service Account Name. Example: *zAdd2Exchange* Press Enter when Finished." | out-file ".\ServiceAccount.txt"
                }
            }

            Else {
                Read-Host "Type in the Sync Service Account Name. Example: *zAdd2Exchange* Press Enter when Finished." | out-file ".\ServiceAccount.txt"
            }


            #Check If Tasks Already Exists

            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists..."
                $confirmation = Read-Host "Would You Like to Update the Current Task? [Y/N]"
                if ($confirmation -eq 'N') {
                    Write-Host "Resuming"
                }

                if ($confirmation -eq 'Y') {
                    Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
                }
            }

            #Check for Powershell File Paths

            $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
            Set-Location $Location

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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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
                    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
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