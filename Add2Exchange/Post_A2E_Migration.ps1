if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #
Do {

    $Title1 = 'Add2Exchange Post Migration Wizard'

    Clear-Host 
    Write-Host "================ $Title1 ================"
    ""
    Write-Host "Please Pick What Best Fits your Scenario"
    ""
    Write-Host "Press '1' for Finishing the Add2Exchange Auto Migration Wizard"
    Write-Host "Press '2' for Add2Exchange is Already Installed and the Console has already been opened Once" 
    Write-Host "Press 'Q' to Quit." -ForegroundColor Red 

    #Migration Method
 
    $input1 = Read-Host "Please Make A Selection" 
    switch ($input1) { 

        #POST Automatic Migration--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '1' { 
            Clear-Host 
            'You chose to Finish the Add2Exchange Auto Migration Wizard'

            # Check for UAC First / Disable UAC
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
            
            #Expanding Setup Files and Database
            Start-Process A2E-Enterprise.exe
            $Setup = Get-ChildItem | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Push-Location $Setup
            Start-Process Powershell ./First_Time_Installer.ps1


            # Migrating Database

            cd..
           
            Write-Host "Migrating Add2Exchange SQL Database. Please Wait....."
            Write-Host "Stopping Add2Exchange SQL Service"
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 5
            Write-Host "Done"

            
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation"
            $Zip = Get-ChildItem | Where-Object { $_.Extension -eq ".zip" } | Select-Object
            

            Rename-Item -Path "$Install\Database\A2E.mdf" -NewName "Clean_A2E.mdf" -ErrorAction SilentlyContinue
            Rename-Item -Path "$Install\Database\A2E_log.LDF" -NewName "Clean_A2E_log.LDF" -ErrorAction SilentlyContinue

            Expand-Archive -Path $Zip -DestinationPath "$Install\Database"
            
            Write-Host "Starting Add2Exchange SQL Service"
            Start-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 5
            Write-Host "Done"

            # Copying Over Rest of Files
            Write-Host "Getting Required Files... Please Wait.."
            
            New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware" -ErrorAction SilentlyContinue
            Copy-Item ".\DidItBetterSoftware*" -Destination "C:\Program Files (x86)\" -Recurse -Force -ErrorAction SilentlyContinue
            REG Import .\License_Info.Reg
            



        }
            

        '2' { 
            Clear-Host 
            'You chose Add2Exchange is Already Installed and the Console has already been opened Once'

            # Check for UAC First / Disable UAC
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
            
            #Expanding Setup Files and Database
            Start-Process A2E-Enterprise.exe
            $Setup = Get-ChildItem | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Push-Location $Setup

            # Registry Favorites & Shortcuts

            Write-Host "Creating Registry Favorites"
            Start-Process Powershell .\Setup\Registry_Favorites.ps1
            
            # Completion and Wrap-Up

            #Removing Outlook Social Connector
            Write-Host "Removing Outlook Social Connector"
            Start-Process Powershell ".\Setup\OSC_Disable.bat"


            # Migrating Database

            cd..
           
            Write-Host "Migrating Add2Exchange SQL Database. Please Wait....."
            Write-Host "Stopping Add2Exchange SQL Service"
            Stop-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 5
            Write-Host "Done"

            
            $Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation"
            $Zip = Get-ChildItem | Where-Object { $_.Extension -eq ".zip" } | Select-Object
            

            Rename-Item -Path "$Install\Database\A2E.mdf" -NewName "Clean_A2E.mdf" -Force -ErrorAction SilentlyContinue
            Rename-Item -Path "$Install\Database\A2E_log.LDF" -NewName "Clean_A2E_log.LDF" -Force -ErrorAction SilentlyContinue

            Expand-Archive -Path $Zip -DestinationPath "$Install\Database" -Force
            
            Write-Host "Starting Add2Exchange SQL Service"
            Start-Service -Name "SQL Server (A2ESQLSERVER)"
            Start-Sleep -s 5
            Write-Host "Done"

            # Copying Over Rest of Files
            Write-Host "Getting Required Files... Please Wait.."
            
            New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware" -ErrorAction SilentlyContinue
            Copy-Item ".\DidItBetterSoftware*" -Destination "C:\Program Files (x86)\" -Recurse -Force -ErrorAction SilentlyContinue
            REG Import .\License_Info.Reg
            

            $wshell = New-Object -ComObject Wscript.Shell

            $answer = $wshell.Popup("Setup is Complete. You can now start the Add2Echange Console", 0, "Done", 0x1)
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