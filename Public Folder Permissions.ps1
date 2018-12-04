if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


$message  = 'Please Pick how you want to connect'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Office 365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Exchange2013-2016'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 2: Quit

if ($decision -eq 2) {

Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}


# Option 1: Office 365


if ($decision -eq 0) {

$error.clear()
Import-Module "MSonline" -ErrorAction SilentlyContinue
If($error){Write-Host "Adding Azure MSonline module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module MSonline -Confirm:$false -WarningAction "Inquire"} 
Else{Write-Host 'Module is installed'}

Import-Module MSOnline

Write-Host "Sign in to Office365 as Tenant Admin"
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session
Import-Module MSOnline
Connect-MsolService -Credential $Cred -ErrorAction "Inquire"


$message  = 'Please Pick what you want to do'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Perm O365'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";


# Option 0: Office 365-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
    do {
    $Identity = read-host "Public Folder Name (Alias)"
    Write-Host "Adding Permissions to Public Folders"
    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
    Write-Host "Done"
    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
    
    } Until ($repeat -eq 'n')
Get-PSSession | Remove-PSSession
Exit
}

# Option 1: Office 365-Removing Public Folder Permissions

if ($decision -eq 1) {
    Write-Host "Getting a list of Public Folders"
    Get-PublicFolder -Identity "\" -Recurse
        do {
        $Identity = read-host "Public Folder Name (Alias)"
        Write-Host "Removing Permissions to Public Folders"
        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
        Write-Host "Done"
        $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        
        } Until ($repeat -eq 'n')
    Get-PSSession | Remove-PSSession
    Exit
}
# Option 2: Office 365-Quit

if ($decision -eq 2) {
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit
}

}
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Option 1: Exchange on Premise


if ($decision -eq 1) {

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Set-ADServerSettings -ViewEntireForest $true

$message  = 'Do you Want to remove or Add Add2Exchange Permissions'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Add Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Exchange Perm'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 0: Exchange on Premise-Adding Public Folder Permissions

#Variables
$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";


# Option 0: Exchange on Premise-Adding Public Folder Permissions

if ($decision -eq 0) {
Write-Host "Getting a list of Public Folders"
Get-PublicFolder -Identity "\" -Recurse
    do {
    $Identity = read-host "Public Folder Name (Alias)"
    Write-Host "Adding Permissions to Public Folders"
    Add-PublicFolderClientPermission -Identity "\$Identity" -User $User -AccessRights Owner -confirm:$false
    Write-Host "Done"
    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
    
    } Until ($repeat -eq 'n')
Get-PSSession | Remove-PSSession
Exit
}

# Option 1: Exchange on Premise-Remove Public Folder Permissions

if ($decision -eq 1) {
    Write-Host "Getting a list of Public Folders"
    Get-PublicFolder -Identity "\" -Recurse
        do {
        $Identity = read-host "Public Folder Name (Alias)"
        Write-Host "Removing Permissions to Public Folders"
        Remove-PublicFolderClientPermission -Identity "\$Identity" -User $User -confirm:$false
        Write-Host "Done"
        $repeat = Read-Host 'Do you want to run it again? [Y/N]'
        
        } Until ($repeat -eq 'n')
    Get-PSSession | Remove-PSSession
    Exit
    }

# Option 2: Exchange on Premise- Quit

if ($decision -eq 2) {
  Write-Host "Quitting"
  Get-PSSession | Remove-PSSession
  Exit
}
}


# End Scripting