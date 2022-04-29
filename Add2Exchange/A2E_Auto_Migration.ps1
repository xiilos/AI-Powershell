if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

#Variables
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue #Current Add2Exchange Installation Path
$BackupDirs = $Install + 'Database\A2E_SQL_Backup'

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
            $A2EBackupDir = "C:\zlibrary\A2E_Backup"
            if ( $(Try { Test-Path $A2EBackupDir.trim() } Catch { $false }) ) {

                Write-Host "A2E_Backup Directory Exists...Resuming"
            }
            Else {
                New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup"
            }

            #Gathering Information about Setup
            Write-Host "Gathering Information about your Setup... Please Wait"
            Copy-Item "$Home\Desktop\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item "$Home\Desktop\Old_Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item "C:\zLibrary\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup\DidItBetterSoftware" -Force
            Copy-Item "C:\Program Files (x86)\DidItBetterSoftware\*" -Destination "C:\zlibrary\A2E_Backup\DidItBetterSoftware\" -Recurse -Force -ErrorAction SilentlyContinue
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\LicenseRegistryInfo" C:\zlibrary\A2E_Backup\License_Info.Reg
            REG EXPORT "HKLM\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" C:\zlibrary\A2E_Backup\Old_Profile_1.Reg
            
            #Upgrade Before Migration
            $Upgrade = Read-Host "Do you want to Upgrade Add2Exchange before migrating? [Y/N]"
            If ($Upgrade -eq 'Y') {
                Start-Process Powershell .\Auto_Upgrade_Add2Exchange.ps1 -Wait
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

            #Backing Up SQL Files
            Write-Host "Backinp Up Add2Exchange SQL Files"
            Start-Process Powershell .\A2E_SQL_Backup.ps1 -Wait
            
            $zip = Get-ChildItem $BackupDirs | Where-Object { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" } | Sort-Object -Property CreationTime -Descending:$True | Select-Object -First 1
            Copy-Item $BackupDirs\$zip -Destination "$A2eBackupDir" -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error.....Cannot Find A2E SQL Backup zip."
                Pause
            }

            Push-Location $env:USERPROFILE
            

            #Download Program Files
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


            
            #Creating Next Steps File

            #Variables

            $NewA2E = Read-Host "What is the Name of your New Add2Exchange Appliance?"
            
            Write-Host "Creating Next Steps File"
            New-Item "C:\zLibrary\A2E_Backup\Next Steps.txt"
            "Next Steps....
            Step 1. Log Into your New Appliance as $env:UserName Account if remaining on the domain
            Step 2. Ensure that $env:UserName is a Local Admin of the new appliance
            Step 3. Copy over the A2E_Backup folder found in C:\zlibrary to $NewA2E
            Step 4. Run the PowerShell File Called Post_A2E_Migration.ps1 in the A2E_Backup Folder on the new Machine" | Out-File -FilePath "C:\zLibrary\A2E_Backup\Next Steps.txt" -Append

            Invoke-Item -Path "C:\zLibrary\A2E_Backup\Next Steps.txt"
            
            #Copy Over Post Migration PowerShell
            $InstallLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
            Push-Location $InstallLocation
            Copy-Item ".\Setup\Post_A2E_Migration.ps1" -Destination "C:\zlibrary\A2E_Backup" -Recurse -ErrorAction SilentlyContinue
            
            $wshell = New-Object -ComObject Wscript.Shell
            $answer = $wshell.Popup("All Files are now Backed up and ready for you to Move them over to the New Add2Exchange Appliance", 0, "Migration Wizard", 0x1)
            if ($answer -eq 2) { Break }

        }

        '2' { 
            Clear-Host 
            'You chose To Just Copy Files Needed and store it in C:\zLibray\A2E_Backup'

            #Creating Landing Zone
            Write-Host "Creating Landing Zone"
            $A2EBackupDir = "C:\zlibrary\A2E_Backup"
            if ( $(Try { Test-Path $A2EBackupDir.trim() } Catch { $false }) ) {

                Write-Host "A2E_Backup Directory Exists...Resuming"
            }
            Else {
                New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup"
            }

            #Gathering Information about Setup
            Write-Host "Gathering Information about your Setup... Please Wait"
            Copy-Item "$Home\Desktop\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item "$Home\Desktop\Old_Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item "C:\zLibrary\Support.txt" -Destination "C:\zlibrary\A2E_Backup" -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -ItemType directory -Path "C:\zlibrary\A2E_Backup\DidItBetterSoftware" -Force
            Copy-Item "C:\Program Files (x86)\DidItBetterSoftware\*" -Destination "C:\zlibrary\A2E_Backup\DidItBetterSoftware\" -Recurse -Force -ErrorAction SilentlyContinue
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
            Start-Process Powershell .\A2E_SQL_Backup.ps1 -Wait
            
            $zip = Get-ChildItem $BackupDirs | Where-Object { $_.Attributes -eq "Archive" -and $_.Extension -eq ".zip" } | Sort-Object -Property CreationTime -Descending:$True | Select-Object -First 1
            Copy-Item $BackupDirs\$zip -Destination "$A2eBackupDir" -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable DB1
            If ($DB1) {
                Write-Host "Error.....Cannot Find A2E SQL Backup zip."
                Pause
            }


            #Download Program Files
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

            #Copy Over Post Migration PowerShell
            $InstallLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
            Push-Location $InstallLocation
            Copy-Item ".\Setup\Post_A2E_Migration.ps1" -Destination "C:\zlibrary\A2E_Backup" -Recurse -ErrorAction SilentlyContinue

            #Creating Next Steps File

            #Variables

            $NewA2E = Read-Host "What is the Name of your New Add2Exchange Appliance?"
            
            Write-Host "Creating Next Steps File"
            New-Item "C:\zLibrary\A2E_Backup\Next Steps.txt"
            "Next Steps....
            Step 1. Log Into your New Appliance as $env:UserName Account if remaining on the domain
            Step 2. Ensure that $env:UserName is a Local Admin of the new appliance
            Step 3. Copy over the A2E_Backup folder found in C:\zlibrary to $NewA2E
            Step 4. Run the PowerShell File Called Post_A2E_Migration.ps1 in the A2E_Backup Folder on the new Machine" | Out-File -FilePath "C:\zLibrary\A2E_Backup\Next Steps.txt" -Append

            Invoke-Item -Path "C:\zLibrary\A2E_Backup\Next Steps.txt"

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