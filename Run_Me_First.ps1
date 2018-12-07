if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Goal
# Create Initial Environment for Add2Exchange Install
# 5 Steps
#Step 1: Disable UAC
#Strp 2: Upgrade .Net and Powershell if needed
#Step 4: Create zLibrary and Download Add2Exchange Software
#Step 5: Account Creation


# Start of Automated Scripting #

# Step 1-----------------------------------------------------------------------------------------------------------------------------------------------------Step 1

# Disable UAC
Write-Host "Disabling UAC In the Registry"
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null
Write-Host "Done"

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
$wshell.Popup("Please Reboot after Installing and run this again",0,"Done",0x1)
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
$wshell.Popup("Please Reboot after Installing",0,"Done",0x1)
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
$wshell.Popup("Please Reboot after Installing",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
    }

Write-Host "Nothing to do"
Write-Host "You Are on the latest version of PowerShell"


#Step 3-----------------------------------------------------------------------------------------------------------------------------------------------------Step 3

#Create zLibrary
$TestPath = "C:\zlibrary"

if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

Write-Host "zLibrary exists...Resuming"
 }
Else {
New-Item -ItemType directory -Path C:\zlibrary
 }

#Download the latest Full Installation
$confirmation = Read-Host "Would you like me to download the latest Add2Exchange? [Y/N]"
    if ($confirmation -eq 'y') {

        Write-Host "Downloading Add2Exchange......"
        Write-Host "This will take a couple of Minutes"
        $url = "ftp://ftp.diditbetter.com/A2E-Enterprise/A2ENewInstall/a2e-enterprise.zip"
        $output = "C:\zlibrary\a2e-enterprise.zip"
        (New-Object System.Net.WebClient).DownloadFile($url, $output)    
        Expand-Archive -Path "C:\zlibrary\a2e-enterprise.zip" -DestinationPath "c:\zlibrary"

}



# Step 4-----------------------------------------------------------------------------------------------------------------------------------------------------Step 4

# Outlook Install 32Bit
If (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application -ErrorAction SilentlyContinue) {
    Write-Output "Outlook Is Installed"
}
Else {
Write-host "We need to install Outlook on this machine"
$confirmation = Read-Host "Would you like me to Install Outlook 365? [Y/N]"
    if ($confirmation -eq 'y') {

Start-Process -filepath "C:\zLibrary\O365Outlook32\OfficeProPlus_x32.msi" -wait

}
}

#Step 5-----------------------------------------------------------------------------------------------------------------------------------------------------Step 5

$message  = 'Have you created an account for Add2Exchange?'
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
Write-host "Please create a new account called zadd2exchange (or any name of your choosing)"
Write-host "Add the new account as a local administrator to this box"
Write-Host "Once done, log off and back on as that account"
Write-Host "Then Run this again"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Log Off and back on as the Service account for Add2Exchange",0,"Done",0x1)
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
    Write-Host "Please logg off and log back on as the new Sync Account"
    Write-Host "Then run this again"
    }
}


Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Initinal Setup is Complete. Please Reboot and run 1 click install",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting