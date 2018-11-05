if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Vairables #

$User = read-host "Enter Sync Service Account name Example: zAdd2Exchange";
$Domain = read-host "Enter your Domain name; If not on a domain please leave this blank"
$Password = read-host "Enter the Account password"

# Start of Automated Scripting #

# Office 365 and on premise Exchange Permissions
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


# Option 0: Office 365


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


# Office 365-Remove&Add Permissions

Write-Host "Removing Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -User $User -accessrights FullAccess -verbose -confirm:$false
Write-Host "Adding Add2Exchange Permissions"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
Write-Host "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_Office365_permissions.txt
Invoke-Item "C:\A2E_Office365_permissions.txt"
Write-Host "Done"
}

# Option 1: Exchange on Premise


if ($decision -eq 1) {


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

Set-ADServerSettings -ViewEntireForest $true

# Option 2: Exchange on Premise-Remove/Add Permissions all

Write-Host 'Removing'
Write-Host "Removing Old zAdd2Exchange Permissions"
Remove-ADPermission -Identity "Exchange Administrative Group (FYDIBOHF23SPDLT)" -User $User -AccessRights ExtendedRight -ExtendedRights "View information store status" -InheritanceType Descendents -Confirm:$false
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights GenericAll -Confirm:$false
Get-Mailbox -Resultsize Unlimited | Remove-mailboxpermission -user $User -accessrights FullAccess -verbose -Confirm:$false
Write-Host "Checking.............................."
Get-MailboxDatabase | Remove-ADPermission -User $User -AccessRights ExtendedRight -ExtendedRights Send-As, Receive-As, ms-Exch-Store-Admin -Confirm:$false
Write-Host "Success....."
Write-Host "Adding Permissions to Users"
Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $User -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
Write-Host "Checking............"
Start-Sleep -s 2
Write-Host "All Done"
Write-Host "Writing Data......"
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} | Select-Object Identity,User, @{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | out-file C:\A2E_permissions.txt
Invoke-Item "C:\A2E_permissions.txt"
Write-Host "Done"
}


# Enable Auto Login
  
Write-Verbose "Enabling Windows AutoLogon"
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 | out-null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $User | out-null
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $Password | out-null
Write-Host "Done"

# Registry Favorites

Write-Host "Creating Registry Favorites"
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Session Manager" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "EnableLUA" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name ".Net Framework" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "OpenDoor Software" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software�" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Add2Exchange"  -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software�\Add2Exchange" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Pendingfilerename" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Session Manager" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "AutoDiscover" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Office" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Group Policy History" -Type string -Value "Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\History" | out-null
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Logon" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" | out-null
Write-Host "Done"

$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Setup is Complete. Please Reboot",0,"Done",0x1)
Get-PSSession | Remove-PSSession
Exit
# End Scripting