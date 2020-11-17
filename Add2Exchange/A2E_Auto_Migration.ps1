if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #
Do {

    $Title1 = 'Add2Exchange Migration Wizard'

    Clear-Host 
    Write-Host "================ $Title1 ================"
    ""
    Write-Host "How would you like to Migrate?"
    ""
    Write-Host "Press '1' for The Migration Wizard"
    Write-Host "Press '2' for Just Copy Files Needed and store it in C:\zLibray\A2E_Backup" 
    Write-Host "Press 'Q' to Quit." -ForegroundColor Red 

    #Migration Method
 
    $input1 = Read-Host "Please Make A Selection" 
    switch ($input1) { 

        #Automatic Migration--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '1' { 
            Clear-Host 
            'You chose The Migration Wizard'

            #Creating Landing Zone
            Write-Host "Creating Landing Zone"
            $TestPath = "C:\zlibrary\A2E_Backup"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

                Write-Host "A2E_Backup Directory Exists...Resuming"
            }
            Else {
                New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup"
            }

            #Gathering Information about Setup
            Write-Host "Gathering Information about your Setup... Please Wait"
            Copy-Item "$Home\Desktop\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            Copy-Item "$Home\Desktop\Old_Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            Copy-Item "C:\zLibrary\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup\DidItBetterSoftware"
            Copy-Item "C:\Program Files (x86)\DidItBetterSoftware\*" -Destination "C:\zlibrary\A2E_Backup\DidItBetterSoftware\" -Recurse -ErrorAction SilentlyContinue
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" C:\zlibrary\A2E_Backup\License_Info.Reg
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" C:\zlibrary\A2E_Backup\Old_Profile_1.Reg
            
            #Upgrade Before Migration
            $Upgrade = Read-Host "Do you want to Upgrade Add2Exchange before Migrating? [Y/N]"
            If ($Upgrade -eq 'Y') {

                #Stop Add2Exchange Service
                Write-Host "Stopping Add2Exchange Service"
                Stop-Service -Name "Add2Exchange Service"
                Start-Sleep -s 2
                Write-Host "Done"

                #Stop The Add2Exchange Agent
                Write-Host "Stopping the Agent. Please Wait."
                Start-Sleep -s 5
                $Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
                if ($Agent) {
                    Write-Host "Waiting for Agent to Exit"
                    Start-Sleep -s 5
                    if (!$Agent.HasExited) {
                        $Agent | Stop-Process -Force
                    }
                }


                #Remove Add2Exchange

                Write-Host "Removing Add2Exchange"
                Write-Host "Please Wait...."
                $Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
                $Program.Uninstall()
                Write-Host "Done"

                #Create zLibrary

                Write-Host "Creating Landing Zone"
                $TestPath = "C:\zlibrary\Add2Exchange Upgrades"
                if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

                    Write-Host "Add2Exchange Upgrades Directory Exists...Resuming"
                }
                Else {
                    New-Item -ItemType directory -Path "C:\zlibrary\Add2Exchange Upgrades"
                }

                #Downloading Add2Exchange

                Write-Host "Downloading Add2Exchange"
                Write-Host "Please Wait......"

                $URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/Upgrades/a2e-enterprise_upgrade.exe"
                $Output = "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"
                $Start_Time = Get-Date

                (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

                Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

                Write-Host "Finished Downloading"

                #Unpacking Add2Exchange

                Write-Host "Unpacking Add2exchange"
                Write-Host "please Wait....."
                Push-Location "c:\zlibrary\Add2Exchange Upgrades"
                Start-Process "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe" -wait
                Start-Sleep -Seconds 2
                Write-Host "Done"

                #Installing Add2Exchange
                Do {
                    Write-Host "Installing Add2Exchange"
                    $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    Push-Location $Location
                    Start-Process -FilePath ".\Add2Exchange_Upgrade.msi" -wait -ErrorAction Inquire -ErrorVariable InstallError;
                    Write-Host "Finished...Upgrade Complete"

                    If ($InstallError) { 
                        Write-Warning -Message "Something Went Wrong with the Install!"
                        Write-Host "Trying The Install Again in 2 Seconds"
                        Start-Sleep -S 2
                    }
                } Until (-not($InstallError))

                #Setting the Service to Delayed Start
                Write-Host "Setting up Add2Exchange Service to Delayed Start"
                sc.exe config "Add2Exchange Service" start= delayed-auto
                Write-Host "Done"

                $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
                Push-Location $Install
                Start-Process ".\Console\Add2Exchange Console.exe"
                Write-Host "Once the Console is Up and Running, Please close it by Hitting File>Exit. Once Finished, click Enter to Continue"
                Pause
    
            }

            If ($Upgrade -eq 'N') {
                Write-Host "Resuming...."
            } 

            #Shutting Down Services

            Write-Host "Stopping Add2Exchange Service"
            Stop-Service -Name "Add2Exchange Service"
            Start-Sleep -s 2
            Write-Host "Done"
            Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service –StartupType Disabled

            #Stop The Add2Exchange Agent
            Write-Host "Stopping the Agent. Please Wait."
            Start-Sleep -s 5
            $Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
            if ($Agent) {
                Write-Host "Waiting for Agent to Exit"
                Start-Sleep -s 5
                if (!$Agent.HasExited) {
                    $Agent | Stop-Process -Force
                }
            }
            Write-Host "Stopping Add2Exchange SQL Service"
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 2
            Write-Host "Done"

            #Backing Up SQL Files
            Write-Host "Backinp Up Add2Exchange SQL Files"
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
            Push-Location $Install
            Copy-Item ".\Database\A2E.mdf" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error.....Cannot Find A2E.mdf Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
                Pause
            }
            Copy-Item ".\Database\A2E_log.LDF" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB2
            If ($DB2) {
                Write-Host "Error.....Cannot Find A2E_Log.LDF Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
                Pause
            }
            
            Push-Location $env:USERPROFILE
            
            #Download Program Files
            $Download = Read-Host "Do you want The Add2Exchange Install Files before Moving? [Y/N]"
            If ($Download -eq 'Y') {
                Write-Host "Downloading Add2Exchange Enterprise"
                Write-Host "Please Wait......"

                $URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/New%20Installs/a2e-enterprise.exe"
                $Output = "C:\zlibrary\A2E_Backup\A2E-Enterprise.exe"
                $Start_Time = Get-Date

                (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

                Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

                Write-Host "Finished Downloading"


                Write-Host "Downloading Recovery and Migration Manager"
                Write-Host "Please Wait......"

                $URL = "ftp://ftp.diditbetter.com/RMM-Enterprise/Upgrades/rmm-enterprise.exe"
                $Output = "C:\zlibrary\A2E_Backup\rmm-enterprise.exe"
                $Start_Time = Get-Date

                (New-Object System.Net.WebClient).DownloadFile($URL, $Output)

                Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

                Write-Host "Finished Downloading"

            }

            If ($Download -eq 'N') {
                Write-Host "Resuming..."
            }
            
            #Creating Next Steps File

            #Variables

            $NewA2E = Read-Host "What is the Name of your New Add2Exchange Appliance?"
            
            Write-Host "Creating Next Steps File"
            New-Item "C:\zLibrary\A2E_Backup\Next Steps.txt"
            "Next Steps....
            Step 1. Log Into your New Appliance as $env:UserName Account if remaing on the domain
            Step 2. Ensure that $env:UserName is a Local Admin of the new appliance
            Step 3. Copy over the A2E_Backup folder found in C:\zlibrary to $NewA2E
            Step 4. Run the PowerShell File Called Post_A2E_Migration.ps1 in the A2E_Backup Folder on the new Machine" | Out-File -FilePath "C:\zLibrary\A2E_Backup\Next Steps.txt" -Append

            Invoke-Item -Path "C:\zLibrary\A2E_Backup\Next Steps.txt"
            
            #Copy Over Post Migration PowerShell
            $InstallLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
            Push-Location $InstallLocation
            Copy-Item ".\Setup\Post_A2E_Migration.ps1" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue 
            
            $wshell = New-Object -ComObject Wscript.Shell
            $answer = $wshell.Popup("All Files are now Backed up and ready for you to Move them over to the New Add2Exchange Appliance", 0, "Migration Wizard", 0x1)
            if ($answer -eq 2) { Break }

        }

        '2' { 
            Clear-Host 
            'You chose To Just Copy Files Needed and store it in C:\zLibray\A2E_Backup'

            #Creating Landing Zone
            Write-Host "Creating Landing Zone"
            $TestPath = "C:\zlibrary\A2E_Backup"
            if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

                Write-Host "A2E_Backup Directory Exists...Resuming"
            }
            Else {
                New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup"
            }

            #Gathering Information about Setup
            Write-Host "Gathering Information about your Setup... Please Wait"
            Copy-Item "$Home\Desktop\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            Copy-Item "$Home\Desktop\Old_Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            Copy-Item "C:\zLibrary\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup\DidItBetterSoftware"
            Copy-Item "C:\Program Files (x86)\DidItBetterSoftware\*" -Destination "C:\zlibrary\A2E_Backup\DidItBetterSoftware\" -Recurse -ErrorAction SilentlyContinue
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" C:\zlibrary\A2E_Backup\License_Info.Reg
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" C:\zlibrary\A2E_Backup\Old_Profile_1.Reg


            #Shuttding Down Services
            Write-Host "Stopping Add2Exchange Service"
            Stop-Service -Name "Add2Exchange Service"
            Start-Sleep -s 2
            Write-Host "Done"
            Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service –StartupType Disabled

            #Stop The Add2Exchange Agent
            Write-Host "Stopping the Agent. Please Wait."
            Start-Sleep -s 5
            $Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
            if ($Agent) {
                Write-Host "Waiting for Agent to Exit"
                Start-Sleep -s 5
                if (!$Agent.HasExited) {
                    $Agent | Stop-Process -Force
                }
            }
            #Stop The Add2Excange SQL Service
            Write-Host "Stopping Add2Exchange SQL Service"
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 2
            Write-Host "Done"

            #Backing Up SQL Files
            Write-Host "Backinp Up Add2Exchange SQL Files"
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
            Push-Location $Install
            Copy-Item ".\Database\A2E.mdf" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error.....Cannot Find A2E.mdf Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
                Pause
            }
            Copy-Item ".\Database\A2E_log.LDF" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB2
            If ($DB2) {
                Write-Host "Error.....Cannot Find A2E_Log.LDF Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
                Pause
            }

            $wshell = New-Object -ComObject Wscript.Shell
            $answer = $wshell.Popup("All Files are now Backed up and ready for you to Move them over to the New Add2Exchange Appliance", 0, "Migration Wizard", 0x1)
            if ($answer -eq 2) { Break }
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