if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Check and Create Stored Credentials

Write-Host "Creating Secure Location"
$TestPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "Secure Location Exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds"
 }

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Create Secure Credentials File

Do {

#Checking Source Tenent or Exchange Admin Username

Write-Host "We will now create your Secure Credentials File"

$TestPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServerUser.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
$confirmation = Read-Host "Exchange/Tenent Admin Username File Exists...Create a New One? [Y/N]"
if ($confirmation -eq 'N') {
Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
Read-Host "Type in the Admin Username to Connect to your Exchange Server. Press Enter when Finished." | out-file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServerUser.txt"
}
}

#Checking Source Tenent or Exchange Admin Password

$TestPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServerPass.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
$confirmation = Read-Host "Exchange/Tenent Admin Password File Exists...Create a New One? [Y/N]"
if ($confirmation -eq 'N') {
Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
Read-Host "Type in the Admin Password or Tenent Admin Password to Connect to your Exchange Server/Office 365. Press Enter when Finished." | out-file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServerPass.txt"
}
}

#Checking Source Exchange Name

$TestPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ExchangeName.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
$confirmation = Read-Host "Exchange Server Name File Exists...Create a New One? [Y/N]"
if ($confirmation -eq 'N') {
Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
Read-Host "Type in your Exchange Server Name. Example: ExchangeServer01 Press Enter when Finished." | out-file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ExchangeName.txt"
}
}

#Checking Source Service Account Name

$TestPath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServiceAccount.txt"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
$confirmation = Read-Host "Service Account Name File Exists...Create a New One? [Y/N]"
if ($confirmation -eq 'N') {
Write-Host "Resuming"
}

if ($confirmation -eq 'Y') {
Read-Host "Type in your Sync Service Account Same. Example: zAdd2Exchange Press Enter when Finished." | out-file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Creds\ServiceAccount.txt"
}
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
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Add Permissions to All Users'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Add Permissions to only Distribution Lists'))
  $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))
  
  $Decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Office 365 All Permissions
if ($Decision -eq 0) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Office365_All_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

#Office 365 Dist. List
if ($Decision -eq 1) {
$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed_Permissions\Office365_Dist_List_Permissions.ps1"'
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

$Trigger= New-ScheduledTaskTrigger -Daily -At 5am
$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file "C:\Timed Permissions\2010_All_Timed_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Option 2: Exchange 2013-2016

if ($choice -eq 2) {

  $Path = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup\Timed Permissions\2013-2016_All_Timed_Permissions.ps1"
  $Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument -ExecutionPolicy Bypass -WindowStyle Hidden $Path
  
  $Trigger =  New-ScheduledTaskTrigger -Daily -At 5am
  Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to users mailboxes"
  }

$repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting