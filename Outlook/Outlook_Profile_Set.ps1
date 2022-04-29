if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Check Outlook Version
$Version = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer" | Select-object -expand '(default)' -ErrorAction SilentlyContinue -ErrorVariable E1
If ($E1) {
  Write-Host "There seems to be an error or conflict in Outlook Versions. Exiting"
  Pause
  Exit
}

#######Outlook V.16#############

If ($Version -eq "Outlook.Application.16") {
#Profile Check
Write-Host "Checking Profile Names"
$Profile = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name ServiceAccount

Write-Host "Add2Exchange Profile in use is $Profile"

Write-Host "Writing $Profile Profile Changes...."

#Address Book
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\9207f3e0a3b11019908b08002b2a56c2" -Name "01023d06" -value ([byte[]](0x00,0x00,0x00,0x00,0xb1,0x84,0xb4,0xea,0xd0,0xab,0xcb,0x41,0xa8,0xd0,0x0d,0xea,0x0c,0x29,0xb0,0x44,0x01,
0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x2f,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "000b3d1c" -value ([byte[]](0x00,0x00))
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "00033d1b" -value ([byte[]](0x01,0x00,0x00,0x00))

#Trust Center
$Privacy = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy"
If ($Privacy -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "Privacy"
}

$Anonymous = Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous"

If ($Anonymous -eq $true){
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesState" -value 2
  $Date = Get-Date -Format 'yyy\-MM\-ddTHH:mm:ssZ'
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesStateTime" -value $Date
}

If ($Anonymous -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy" -Name "SettingsStore"
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore" -Name "Anonymous"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesState" -value 2
  $Date = Get-Date -Format 'yyy\-MM\-ddTHH:mm:ssZ'
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesStateTime" -value $Date
}


$Profile1 = Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e"

If ($Profile1 -eq $true){
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e" -Name "00030354" -value ([byte[]](0x00,0x00,0x00,0x00))
}

If ($Profile1 -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile" -Name "c02ebc5353d9cd11975200aa004ae40e"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e" -Name "00030354" -value ([byte[]](0x00,0x00,0x00,0x00))
}

$Security = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security"
If ($Security -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "Security"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "InitEncrypt" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "InitSign" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "SharedFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "PublicFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "Level" -value 1

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name "DisableAttachmentPreviewing" -value 1

$Mail = Test-Path -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail"
If ($Mail -eq $true) {
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSafeZone" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSpecificSenders" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockRSS" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSTS" -value 0
}


If ($Mail -eq $false) {
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options" -Name "Mail"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSafeZone" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSpecificSenders" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockRSS" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail" -Name "UnblockSTS" -value 0
  }

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerDataOptInReason" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerDataOptIn" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "SendCustomerData" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "UpdateReliabilityData" -value 0

$Helperviewer = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer"
If ($Helperviewer -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "HelpViewer"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\HelpViewer" -Name "UseOnlineContent" -value 2

$PTWatson = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson"
If ($PTWatson -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common" -Name "PTWatson"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\PTWatson" -Name "PTWOptIn" -value 3

$ResearchOptions = Test-Path "HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options"
If ($ResearchOptions -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Research" -Name "Options"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Research\Options" -Name "DiscoveryNeedOptIn" -value 0

$Security = Test-Path -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security"

If ($Security -eq $true){
  Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security" -Name "ObjectModelGuard" -value 2
}

If ($Security -eq $false){
  New-Item -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook" -Name "Security"
  New-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security" -Name "ObjectModelGuard" -value 2
}


#Outlook Social Connector
Write-Host "Disabling Outlook Social Connector"
Start-Process Powershell .\OSC_Disable.bat
Write-Host "Done"

#Disable Outlook Updates
Write-Host "Disabling Outlook Updates"

$Val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled"

if($val.UpdatesEnabled -eq $True)

{
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled" -value False
Write-Host "Outlook Updates are now Disabled!"

}

Else {

  Write-Host "Outlook Updates Already Disabled!"

}

#Disable Outlook popups

Write-Host "Disabling Teaching Callouts"
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "AutoSaveTottleOnWord" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "MeetingAllowForwardTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "AutoSaveFirstSaveWord" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "CommingSoonTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "AutocreateTeachingCallout_MoreLocations" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "Search.TopResults" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "UseTighterSpacingTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "SLRToggleReplaceTeachingCalloutID" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "DataVisualizerRibbonTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "ExportToWordProcessTabTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Common\TeachingCallouts" -Name "PreviewPlaceUpdate" -value 0 -ErrorAction SilentlyContinue

} 

#######Outlook V.15#############

If ($Version -eq "Outlook.Application.15") {
  #Profile Check
Write-Host "Checking Profile Names"
$Profile = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name ServiceAccount

Write-Host "Add2Exchange Profile in use is $Profile"

Write-Host "Writing $Profile Profile Changes...."

#Address Book
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\9207f3e0a3b11019908b08002b2a56c2" -Name "01023d06" -value ([byte[]](0x00,0x00,0x00,0x00,0xb1,0x84,0xb4,0xea,0xd0,0xab,0xcb,0x41,0xa8,0xd0,0x0d,0xea,0x0c,0x29,0xb0,0x44,0x01,
0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x2f,0x00))

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "000b3d1c" -value ([byte[]](0x00,0x00))
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\0a0d020000000000c000000000000046" -Name "00033d1b" -value ([byte[]](0x01,0x00,0x00,0x00))

#Trust Center
$Privacy = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy"
If ($Privacy -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "Privacy"
}

$Anonymous = Test-Path -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore\Anonymous"

If ($Anonymous -eq $true){
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesState" -value 2
  $Date = Get-Date -Format 'yyy\-MM\-ddTHH:mm:ssZ'
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesStateTime" -value $Date
}

If ($Anonymous -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy" -Name "SettingsStore"
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore" -Name "Anonymous"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesState" -value 2
  $Date = Get-Date -Format 'yyy\-MM\-ddTHH:mm:ssZ'
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Privacy\SettingsStore\Anonymous" -Name "ControllerConnectedServicesStateTime" -value $Date
}


$Profile1 = Test-Path -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e"

If ($Profile1 -eq $true){
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e" -Name "00030354" -value ([byte[]](0x00,0x00,0x00,0x00))
}

If ($Profile1 -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile" -Name "c02ebc5353d9cd11975200aa004ae40e"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles\$Profile\c02ebc5353d9cd11975200aa004ae40e" -Name "00030354" -value ([byte[]](0x00,0x00,0x00,0x00))
}

$Security = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security"
If ($Security -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook" -Name "Security"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security" -Name "InitEncrypt" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security" -Name "InitSign" -value 2
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security" -Name "SharedFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security" -Name "PublicFolderScript" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Security" -Name "Level" -value 1

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Preferences" -Name "DisableAttachmentPreviewing" -value 1

$Mail = Test-Path -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail"
If ($Mail -eq $true) {
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSafeZone" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSpecificSenders" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockRSS" -value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSTS" -value 0
}


If ($Mail -eq $false) {
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options" -Name "Mail"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSafeZone" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSpecificSenders" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockRSS" -value 0
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail" -Name "UnblockSTS" -value 0
  }

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "SendCustomerDataOptInReason" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "SendCustomerDataOptIn" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "SendCustomerData" -value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "UpdateReliabilityData" -value 0

$Helperviewer = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Common\HelpViewer"
If ($Helperviewer -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "HelpViewer"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\HelpViewer" -Name "UseOnlineContent" -value 2

$PTWatson = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Common\PTWatson"
If ($PTWatson -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common" -Name "PTWatson"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\PTWatson" -Name "PTWOptIn" -value 3

$ResearchOptions = Test-Path "HKCU:\Software\Microsoft\Office\15.0\Common\Research\Options"
If ($ResearchOptions -eq $false){
  New-Item -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Research" -Name "Options"
}

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Research\Options" -Name "DiscoveryNeedOptIn" -value 0

$Security = Test-Path -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\15.0\Outlook\Security"

If ($Security -eq $true){
  Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\15.0\Outlook\Security" -Name "ObjectModelGuard" -value 2
}

If ($Security -eq $false){
  New-Item -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\15.0\Outlook" -Name "Security"
  New-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\15.0\Outlook\Security" -Name "ObjectModelGuard" -value 2
}


#Outlook Social Connector
Write-Host "Disabling Outlook Social Connector"
Start-Process Powershell .\OSC_Disable.bat
Write-Host "Done"


#Disable Outlook Updates
Write-Host "Disabling Outlook Updates"

$Val = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled"

if($val.UpdatesEnabled -eq $True)

{
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled" -value False
Write-Host "Outlook Updates are now Disabled!"

}

Else {

  Write-Host "Outlook Updates Already Disabled!"

}

#Disable Outlook popups

Write-Host "Disabling Teaching Callouts"
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "AutoSaveTottleOnWord" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "MeetingAllowForwardTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "AutoSaveFirstSaveWord" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "CommingSoonTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "AutocreateTeachingCallout_MoreLocations" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "Search.TopResults" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "UseTighterSpacingTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "SLRToggleReplaceTeachingCalloutID" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "DataVisualizerRibbonTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "ExportToWordProcessTabTeachingCallout" -value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\TeachingCallouts" -Name "PreviewPlaceUpdate" -value 0 -ErrorAction SilentlyContinue



}

#######Outlook Not Supported#############

If ($Version -eq "Outlook.Application.14") { 
  Write-Host "This version of Outlook is not supported"
}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting