if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

# Disable UAC From Registry

Write-Verbose "Disabling UAC In the Registry"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null

 # Setting Time Zone
 Write-Verbose "Setting the Time Zone"
Set-TimeZone -Name "Central Standard Time"
    
# Rename Computer

Write-Verbose "Renaming Computer"
$Compname = read-host "Enter the new name of this Computer"
Rename-Computer -NewName "$CompName"


# Rename HDD's
Write-Verbose "Renaming HDD's"
Get-Volume -DriveLetter C | Set-Volume -NewFileSystemLabel "MAIN OS" -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    write-warning -message "Something went wrong renaming C";

} 

Get-Volume -DriveLetter D | Set-Volume -NewFileSystemLabel "DATA" -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    write-warning -message "Something went wrong renaming D!";

} 


# File Explorer Options "Views"

Write-Verbose "Changing File Explorer Views"
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name Hidden -Value 1 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name AutoCheckSelect -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name showsuperhidden -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideDrivesWithNoMedia -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideFileExt -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarGlomLevel -Value 2 | out-null
    Set-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer' -Name EnableAutoTray -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name DontPrettyPath -Value 1 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name HideMergeConflicts -Value 0 | out-null
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState' -Name FullPath -Value 1 | out-null
    Stop-Process -processname explorer




# Disable Drive indexing
Write-Verbose "Disabling Drive Indexing"
function Disable-Indexing {
    Param($Drive)
    $obj = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='$Drive'"
    $indexing = $obj.IndexingEnabled
    if("$indexing" -eq $True){
        write-host "Disabling indexing of drive $Drive"
        $obj | Set-WmiInstance -Arguments @{IndexingEnabled=$False} | Out-Null
    }
}

#Use:
#Disable Drive Indexing on C:\ and D:\

Disable-Indexing "C:"
Disable-Indexing "D:"


# Disabled Services

Write-Verbose "Disabling Services"
Stop-Service WSearch
Set-Service WSearch -StartupType Disabled

Stop-Service WpcMonSvc
Set-Service WpcMonSvc -StartupType Disabled

Stop-Service wercplsupport
Set-Service wercplsupport -StartupType Disabled

Stop-Service RetailDemo
Set-Service RetailDemo -StartupType Disabled

Stop-Service SDRSVC
Set-Service SDRSVC -StartupType Disabled

Stop-Service WerSvc
Set-Service WerSvc -StartupType Disabled

Stop-Service XboxGipSvc
Set-Service XboxGipSvc -StartupType Disabled

Stop-Service XblAuthManager
Set-Service XblAuthManager -StartupType Disabled

Stop-Service XblGameSave
Set-Service XblGameSave -StartupType Disabled

Stop-Service XboxNetApiSvc
Set-Service XboxNetApiSvc -StartupType Disabled

# Control Panel Options
Write-Verbose "Setting Mouse Speed"
Set-ItemProperty -Path 'HKCU:\Control Panel\Mouse' -Name MouseSensitivity -Value 20 | out-null

Write-Verbose "Allowing RDC"
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name AllowRemoteRPC -Value 0 | out-null


Write-Verbose "Allowing RDC"
Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0 | out-null
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"

Write-Verbose "Disable Syetm Restore"
Disable-ComputerRestore "C:\", "D:\"

Write-Verbose "Setting Visual Performance"
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' -Name VisualFXSetting -Value 2 | out-null

# Windows Apps Cleanup

Write-Verbose "Make Sure to Enable Cloud Content GPO"


Write-Verbose "Removing Windows Apps"
Get-AppxPackage -AllUsers | where-object {$_.name –notlike "*windows.photos"} | where-object {$_.name –notlike "*store*"} | where-object {$_.name –notlike "*calculator*"} | where-object {$_.name –notlike "*sticky*"} | where-object {$_.name –notlike "*soundrecorder*"} | where-object {$_.name –notlike "*mspaint*"} | where-object {$_.name –notlike "*screensketch*"} | Remove-AppxPackage -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    write-warning -message "Cannot Remove this App";

} 
Get-AppxPackage | Select-Object Name, PackageFullName >"$env:userprofile\Desktop\myapps.txt"


# Windows Settings Cleanup

