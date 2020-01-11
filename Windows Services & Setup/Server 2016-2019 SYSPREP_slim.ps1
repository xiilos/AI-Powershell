if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}


# Disable UAC From Registry
Write-Host "Disabling UAC In the Registry"
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | Out-Null

# Setting Time Zone
Write-Host "Setting the Time Zone CST"
Set-TimeZone -Name "Central Standard Time"

    

# Rename HDD's
Write-Host "Renaming HDD's"
Get-Volume -DriveLetter C | Set-Volume -NewFileSystemLabel "MAIN OS" -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    Write-Warning -message "Something went wrong renaming C";

} 

Get-Volume -DriveLetter D | Set-Volume -NewFileSystemLabel "DATA" -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    Write-Warning -message "Drive D:\ Does not Exist";

} 


# File Explorer Options "Views"

Write-Host "Changing File Explorer Views"
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name Hidden -Value 1 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name AutoCheckSelect -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name showsuperhidden -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideDrivesWithNoMedia -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideFileExt -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarGlomLevel -Value 2 | Out-Null
Set-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer' -Name EnableAutoTray -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name DontPrettyPath -Value 1 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideMergeConflicts -Value 0 | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState' -Name FullPath -Value 1 | Out-Null
Stop-Process -processname explorer


# Disable Drive indexing

Write-Host "Disabling Drive Indexing"
function Disable-Indexing {
    Param($Drive)
    $obj = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='$Drive'"
    $indexing = $obj.IndexingEnabled
    if ("$indexing" -eq $True) {
        Write-Host "Disabling indexing of drive $Drive"
        $obj | Set-WmiInstance -Arguments @{IndexingEnabled = $False } | Out-Null
    }
}

#Disable Drive Indexing on C:\ and D:\

Disable-Indexing "C:"
Disable-Indexing "D:" -ErrorAction SilentlyContinue


# Control Panel Options

Write-Host "Setting Mouse Speed"
Set-ItemProperty -Path 'HKCU:\Control Panel\Mouse' -Name MouseSensitivity -Value 20 | Out-Null

Write-Host "Removing Remote Assistance"
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance" -Name "fAllowToGetHelp" -Type DWord -Value 0 -ErrorAction SilentlyContinue

Write-Host "Allowing RDC"
Write-Host "Enabling Remote Desktop w/o Network Level Authentication..."
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Type DWord -Value 0
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -Type DWord -Value 0
Enable-NetFirewallRule -Name "RemoteDesktop*"

Write-Host "Disable Syetm Restore"
Disable-ComputerRestore "C:\" -ErrorAction SilentlyContinue
Disable-ComputerRestore "D:\" -ErrorAction SilentlyContinue -ErrorVariable ProcessError;
If ($ProcessError) {

    Write-Warning -message "D:\ Does not Exist";
}

Write-Host "Setting Visual Performance"
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -Name VisualFXSetting -Value 2 | Out-Null

Write-Host "Set Control Panel Items to Large View"
If (!(Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel")) {
    New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" | Out-Null
}
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" -Name "StartupPage" -Type DWord -Value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel" -Name "AllItemsIconView" -Type DWord -Value 0


Write-Host "Explorer view from Quick Access to This PC"
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Type DWord -Value 1



# Setup Network to Private and File Sharing

Write-Host "Setting current network profile to private..."
Set-NetConnectionProfile -NetworkCategory Private


# Disable Windows Defender Cloud Submission

Write-Host "Disabling Windows Defender Cloud..."
If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet")) {
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Force | Out-Null
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SpynetReporting" -Type DWord -Value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SubmitSamplesConsent" -Type DWord -Value 2


#Set Internet Explorer Settings

$Path = 'HKCU:\Software\Microsoft\Internet Explorer\Main\'
$Name = 'start page'
$Value = 'About:none'
Set-Itemproperty -Path $Path -Name $Name -Value $Value


Write-Output "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting