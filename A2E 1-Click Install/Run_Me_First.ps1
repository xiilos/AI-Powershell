if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Goal
# Create Initial Environment for Add2Exchange Install
# 7 Steps

#Step 1: Account Creation
#Step 2: Upgrade .Net and Powershell if needed
#Step 3: Create zLibrary and Download Add2Exchange Software
#Step 4: Install Outlook and Setup Profile
#Step 5: Mailbox Creation
#Step 6: Create a Mail Profile
#Step 7: Disable UAC


# Start of Automated Scripting #

#Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1
# Account Creation

$message  = 'Have you created an account for Add2Exchange to install under?'
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
Log off and back on as the new Sync Service account and run this again before you Proceed.",0,"Account Creation",0x1)
if($answer -eq 2){Break}

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
Log off and back on as the new Sync Service account and run this again before proceeding.",0,"Account Creation",0x1)
if($answer -eq 2){Break}

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

if (($installed -ne 1) -or ($release -lt 378389))
{
Write-Host "We need to download .Net 4.5.2"
Write-Host "Downloading"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}

$url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
$output = "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
(New-Object System.Net.WebClient).DownloadFile($url, $output)    
Start-Process -FilePath "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
if($answer -eq 2){Break}
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


#Check Operating Sysetm

$BuildVersion = [System.Environment]::OSVersion.Version


#OS is 10+
if($BuildVersion.Major -like '10')
    {
        Write-Host "WMF 5.1 is not supported for Windows 10 and above"
    }

#OS is 7
if($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1')
    {
        
Write-Host "Downloading WMF 5.1 for 7+"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}
$url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
$output = "C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Process -FilePath 'C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu' -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
if($answer -eq 2){Break}
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

#OS is 8
elseif($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3')
    {
        
Write-Host "Downloading WMF 5.1 for 8+"
$Directory = "C:\PowerShell"

if ( -Not (Test-Path $Directory.trim() ))
{
    New-Item -ItemType directory -Path C:\PowerShell
}
$url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
$output = "C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Start-Process -FilePath 'C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu' -wait
Write-Host "Download Complete"
$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
if($answer -eq 2){Break}
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

 Copy-Item -Path ".\support.txt" -Destination "$home\Desktop\Support.txt"
 Copy-Item -Path ".\Setup\PermissionsOnPremOrO365Combined.ps1" -Destination "$home\Desktop\PermissionsOnPremOrO365Combined.ps1"

 $WshShell = New-Object -comObject WScript.Shell
 $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\zLibrary.lnk")
 $Shortcut.TargetPath = "C:\zlibrary"
 $Shortcut.Save()

#Download the latest Full Installation

##IN-PRODUCTION##

# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4

# Outlook Install 32Bit

If (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application -ErrorAction SilentlyContinue) {
    Write-Output "Outlook Is Installed"
}
Else {
Write-host "We need to install Outlook on this machine"
$confirmation = Read-Host "Would you like me to Install Outlook 365? [Y/N]"
if ($confirmation -eq 'y') {

Push-Location -Path ".\O365Outlook32"
.\setup.exe /configure Office365x32bit.xml

}

if ($confirmation -eq 'n') {

$wshell = New-Object -ComObject Wscript.Shell
$answer = $wshell.Popup("If you choose to install your own Outlook; ensure that it is Outlook 2016 and above, and 32bit Only.
When Done, click OK",0,"Outlook Install",0x1)
if($answer -eq 2){Break}

    }
}


# Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5

# Mailbox Creation

$confirmation = Read-Host "Are you on Office 365 or Exchange on Premise [O/E]"
if ($confirmation -eq 'O') {
Write-host "Make sure to create a mailbox for the sync service account and add an *E* License to it"
$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("When this is Done, click OK to Continue",0,"Create a Mailbox",0x1)
if($answer -eq 2){Break}
}

if ($confirmation -eq 'E') {
Write-Host "Create a Mailboix for the new Sync Service account in Exchange"
$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("When this is Done, Click OK to Continue",0,"Create a Mailbox",0x1)
if($answer -eq 2){Break}
}

# Step 6-----------------------------------------------------------------------------------------------------------------------------------------------------Step 6

# Mail Profile

$wshell = New-Object -ComObject Wscript.Shell

$answer = $wshell.Popup("The next step is to Create a Profile for your new account. Open Control panel and go to Mail. Create a new profile and follow through the steps that pertain to your Organization. 
Note* Make sure you do not have Cache checked. When this is finished click OK to Continue",0,"Creating an Outlook Profile",0x1)
if($answer -eq 2){Break}

# Step 7-----------------------------------------------------------------------------------------------------------------------------------------------------Step 7

# Disable UAC
Write-Host "Checking UAC"

$Val = Get-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA"

if($val.EnableLUA -ne 0)

{
Set-ItemProperty -Path "HKLM:Software\Microsoft\Windows\Currentversion\Policies\System" -Name "EnableLUA" -value 0
Write-Host "Done"
Write-Host "UAC is now Disabled"
$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Initinal Setup is Complete. Please Reboot and run 1-click install",0,"Done",0x1)
if($answer -eq 2){Break}
  
}

Else {

Write-Host "UAC is already Disabled"
Write-Host "Resuming"
Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Initinal Setup is Complete. Please run 1-click install",0,"Done",0x1)
if($answer -eq 2){Break}
    
}

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting