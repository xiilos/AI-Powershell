if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Write-Host "Checking Profile Names"
$Profile = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name ServiceAccount

Write-Host "Add2Exchange Profile in use is $Profile"

Write-Host "Writing $Profile Profile Changes...."

#Address Book
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\9207f3e0a3b11019908b08002b2a56c2" -Name "01023d06" -value ([byte[]](0x00,0x00,0x00,0x00,0xb1,0x84,0xb4,0xea,0xd0,0xab,0xcb,0x41,0xa8,0xd0,0x0d,0xea,0x0c,0x29,0xb0,0x44,0x01,
0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x2f,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "000b3d1c" -value ([byte[]](0x00,0x00))
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "00033d1b" -value ([byte[]](0x00,0x00,0x00,0x00))

#Trust Center
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesState" -value 2

$Date = Get-Date -Format 'yyy\-MM\-ddTHH:mm:ssZ'
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesStateTime" -value $Date

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e" -Name "00030354" -value ([byte[]](0x00,0x00,0x00,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "InitEncrypt" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "InitSign" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "SharedFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "PublicFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "Level" -value 1

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name "DisableAttachmentPreviewing" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSafeZone" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSpecificSenders" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockRSS" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSTS" -value 0

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerDataOptInReason" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerDataOptIn" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerData" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "UpdateReliabilityData" -value 0

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer" -Name "UseOnlineContent" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson" -Name "PTWOptIn" -value 3
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options" -Name "DiscoveryNeedOptIn" -value 0

Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security" -Name "ObjectModelGuard" -value 2

#Outlook Social Connector
Start-Process Powershell .\OSC_Disable.bat

Write-Host "Done"
Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting