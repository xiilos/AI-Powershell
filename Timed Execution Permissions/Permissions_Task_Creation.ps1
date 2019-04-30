if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Check for Powershell File Paths

<#$Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
Set-Location $Location

#Check and Create Stored Credentials

$TestPath = ".\Setup\Timed Permissions\Creds"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    Write-Host "Secure Location Exists...Resuming"
}
Else {
    Write-Host "Creating Secure Location"
    New-Item -ItemType directory -Path ".\Setup\Timed Permissions\Creds"
}
#>

#Check and Create Stored Credentials

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

#Create Secure Credentials File

Do {
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
            Read-Host "Type in your Global Admin or Exchange Admin Username. Leave Blank for None. Press Enter when Finished." | out-file ".\ServerUser.txt"
        }

    }

    Else {
        Read-Host "Type in your Global Admin or Exchange Admin Username. Leave Blank for None. Press Enter when Finished." | out-file ".\ServerUser.txt"
    }
   
    #Checking Source Tenent or Exchange Admin Password

    $TestPath = ".\ServerPass.txt"
    if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
        Write-Host "Exchange/Tenent Admin Password File Exists..."
        Write-Host "Last Updated:" -ForegroundColor Green
        Get-Item "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\ServerPass.txt" | ForEach-Object {$_.LastWriteTime}
        ""
        $confirmation = Read-Host "Would You Like to Update the Current File? [Y/N]"
        if ($confirmation -eq 'N') {
            Write-Host "Resuming"
        }
   
        if ($confirmation -eq 'Y') {
            Read-Host "Type in your Global Admin or Exchange Admin Password. Press Enter when Finished." -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
        }

    }

    Else {
        Read-Host "Type in your Global Admin or Exchange Admin Password. Press Enter when Finished." -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"
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
            Read-Host "Type in the Sync Service Account Same. Example: zAdd2Exchange Press Enter when Finished." | out-file ".\ServiceAccount.txt"
        }
    }

    Else {
        Read-Host "Type in the Sync Service Account Same. Example: zAdd2Exchange Press Enter when Finished." | out-file ".\ServiceAccount.txt"
    }


    #Check If Tasks Already Exists

    if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction Ignore) {
        Write-Host "Add2Exchange Permissions Task Already Exists..."
        $confirmation = Read-Host "Would You Like to Update the Current Task? [Y/N]"
        if ($confirmation -eq 'N') {
            Write-Host "Resuming"
        }
   
        if ($confirmation -eq 'Y') {
            Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
        }
    }

    Else { Write-Host "How are you logging on?"}

    #Check for Powershell File Paths

    $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
    Set-Location $Location

    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    #Connection Method

    $message = 'Please Pick how you want to connect'
    $question = 'Pick one of the following from below'

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&MExchange 2010'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange 2013-2016'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

    $choice = $Host.UI.PromptForChoice($message, $question, $choices, 3)

    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Option 0: Office 365

    if ($choice -eq 0) {

        $message = 'Please Pick how you want to connect'
        $question = 'Pick one of the following from below'
  
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
        $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

        #Office 365 All Permissions
        if ($Decision -eq 0) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }


        #Office 365 Dist. List
        if ($Decision -eq 1) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dist_List_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }

        #Quitting Office 365
        if ($Decision -eq 2) {
            Write-Host "Quitting"
            Get-PSSession | Remove-PSSession
            Exit

        }
    }
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Option 1: Exchange 2010

    if ($choice -eq 1) {

        $message = 'Please Pick how you want to connect'
        $question = 'Pick one of the following from below'
  
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
        $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

        #Exchange 2010 All Permissions
        if ($Decision -eq 0) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010_All_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }

        #Exchange 2010 Dist. List
        if ($Decision -eq 1) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010_Dist_List_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }

        #Quitting Exchange 2010
        if ($Decision -eq 2) {
            Write-Host "Quitting"
            Get-PSSession | Remove-PSSession
            Exit

        }
    }

    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Option 2: Exchange 2013-2016

    if ($choice -eq 2) {

        $message = 'Please Pick how you want to connect'
        $question = 'Pick one of the following from below'
  
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
        $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

        #Exchange 2013-2016 All Permissions
        if ($Decision -eq 0) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2013-2016_All_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }

        #Exchange 2013-2016 Dist. List
        if ($Decision -eq 1) {
            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2013-2016_Dist_List_Permissions.ps1"'
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
        }

        #Quitting Exchange 2013-2016
        if ($Decision -eq 2) {
            Write-Host "Quitting"
            Get-PSSession | Remove-PSSession
            Exit

        }
    }

    #--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting