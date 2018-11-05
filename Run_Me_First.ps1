if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Start of Automated Scripting #

# Disable UAC
Write-Host "Disabling UAC In the Registry"
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null
Write-Host "Done"


#Create zLibrary
New-Item -ItemType directory -Path C:\zlibrary

# Powershell Update
# Check if .Net 4.5 or above is installed

$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389))
{
    Write-Host "We need to download .Net 4.5.2"
    Write-Host "Downloading"
    $url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    $output = "c:\zlibrary\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Write-Host "Download Complete"
    
    Invoke-item -Path "c:\zlibrary\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Please Reboot After Install and run this again",0,"Done",0x1)
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
$url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
$output = "c:\zlibrary\Win7AndW2K8R2-KB3191566-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Write-Host "Download Complete"
    
    Invoke-Item -Path 'c:\zlibrary\Win7AndW2K8R2-KB3191566-x64.msu'
    $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Please Reboot after install and run this again",0,"Done",0x1)
    }

#OS is 8
elseif($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3')
    {
        
Write-Host "Downloading WMF 5.1 for 8+"
$url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
$output = "c:\zlibrary\Win8.1AndW2K12R2-KB3191564-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Write-Host "Download Complete"
    
    Invoke-Item -Path 'c:\zlibrary\Win8.1AndW2K12R2-KB3191564-x64.msu'
    $wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Please Reboot after install and run this again",0,"Done",0x1)
    }

# Outlook 365 / 2016 Install 32Bit
If (Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application -ErrorAction SilentlyContinue) {
    Write-Output "Outlook Is Installed"
}
Else {
Write-host "We need to install Outlook on this machine"
$confirmation = Read-Host "Would you like me to Install Outlook 365? [Y/N]"
    if ($confirmation -eq 'y') {


}


}

$message  = 'Have you created an account for Add2Exchange'
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
}

if ($confirmation -eq 'n') {
Write-Host "Lets create a new user for Add2Exchange"
$Username = Read-Host "Username for new account"
$Fullname = Read-host "User Full Name"    
$password = Read-Host "Password for new account"
$description = read-host "Description of this account."

    New-LocalUser $Username -Password $Password -FullName $Fullname -Description $description
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