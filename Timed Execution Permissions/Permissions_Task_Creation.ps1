if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Check for Powershell File Paths

$Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
Set-Location $Location

#Check and Create Stored Credentials

Write-Host "Creating Secure Location"
$TestPath = ".\Setup\Timed Permissions\Creds"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
Write-Host "Secure Location Exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path ".\Setup\Timed Permissions\Creds"
 }

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Create Secure Credentials File

Do {
Write-Host "We will now create your Secure Credential Files"
#Checking Source Tenent or Exchange Admin Username

$TestPath = ".\Setup\Timed Permissions\Creds\ServerUser.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  Write-Host "Exchange/Tenent Admin Username File Exists...Resuming"
   }

  Else {
  Read-Host "Type in a Username to Connect to the Exchange Server. Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ServerUser.txt"
   }
   $confirmation = Read-Host "Do it Again? or Create a New One? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
   Read-Host "Type in a Username to Connect to the Exchange Server. Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ServerUser.txt"
   }

#Checking Source Tenent or Exchange Admin Password

$TestPath = ".\Setup\Timed Permissions\Creds\ServerPass.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  Write-Host "Exchange/Tenent Admin Password File Exists...Resuming"
   }

  Else {
  Read-Host "Type in the Admin/Tenent Admin Password to Connect to your Exchange Server/Office 365. Press Enter when Finished." -assecurestring | convertfrom-securestring | out-file ".\Setup\Timed Permissions\Creds\ServerPass.txt"
   }
   $confirmation = Read-Host "Do it Again? or Create a New One? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
    Read-Host "Type in the Admin/Tenent Admin Password to Connect to your Exchange Server/Office 365. Press Enter when Finished." -assecurestring | convertfrom-securestring | out-file ".\Setup\Timed Permissions\Creds\ServerPass.txt"
   }

#Checking Source Exchange Name

$TestPath = ".\Setup\Timed Permissions\Creds\ExchangeName.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  Write-Host "Exchange Server Name File Exists...Resuming"
   }

  Else {
  Read-Host "Type in your Exchange Server Name. Example: ExchangeServer01 Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ExchangeName.txt"
   }
   $confirmation = Read-Host "Do it Again? or Create a New One? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
   Read-Host "Type in your Exchange Server Name. Example: ExchangeServer01 Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ExchangeName.txt"
   }

#Checking Source Distribution List Name

$TestPath = ".\Setup\Timed Permissions\Creds\DistributionName.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  Write-Host "Distribution List Name File Exists...Resuming"
   }

  Else {
  Read-Host "Type in your Distribution List Name. Example: zFirmCalendarSync Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\DistributionName.txt"
   }
   $confirmation = Read-Host "Do it Again? or Create a New One? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
   Read-Host "Type in your Distribution List Name. Example: zFirmCalendarSync Leave Blank for None. Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\DistributionName.txt"
   }

#Checking Source Service Account Name

$TestPath = ".\Setup\Timed Permissions\Creds\ServiceAccount.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
  Write-Host "Service Account Name File Exists...Resuming"
   }

  Else {
  Read-Host "Type in your Sync Service Account Same. Example: zAdd2Exchange Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ServiceAccount.txt"
   }
   $confirmation = Read-Host "Do it Again? or Create a New One? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
   Read-Host "Type in your Sync Service Account Same. Example: zAdd2Exchange Press Enter when Finished." | out-file ".\Setup\Timed Permissions\Creds\ServiceAccount.txt"
   }


#Check If Tasks Already Exists

if(Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction Ignore) 
{
  Write-Host "Add2Exchange Permissions Task Already Exists..."
}
 else { Write-Host "Add2Exchange Permissions Task Does not Exist Yet... Resuming..."}

 $confirmation = Read-Host "Create a New Task and Overwrite Old Add2Exchange Task? [Y/N]"
   if ($confirmation -eq 'N') {
   Write-Host "Resuming"
   }
   
   if ($confirmation -eq 'Y') {
    Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
   } 
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Connection Method

$message  = 'Please Pick how you want to connect'
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

  $message  = 'Please Pick how you want to connect'
  $question = 'Pick one of the following from below'
  
  $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
  $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Office 365 All Permissions
if ($Decision -eq 0) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file .\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

#Office 365 Dist. List
if ($Decision -eq 1) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dist_List_Permissions.ps1"'
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

  $message  = 'Please Pick how you want to connect'
  $question = 'Pick one of the following from below'
  
  $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
  $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Exchange 2010 All Permissions
if ($Decision -eq 0) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010_All_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

#Exchange 2010 Dist. List
if ($Decision -eq 1) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010_Dist_List_Permissions.ps1"'
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

  $message  = 'Please Pick how you want to connect'
  $question = 'Pick one of the following from below'
  
  $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1-Add Permissions to All Users'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2-Add Permissions to only Distribution Lists'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
  $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Exchange 2013-2016 All Permissions
if ($Decision -eq 0) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2013-2016_All_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

#Exchange 2013-2016 Dist. List
if ($Decision -eq 1) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2013-2016_Dist_List_Permissions.ps1"'
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