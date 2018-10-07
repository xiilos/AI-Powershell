if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

# Disable UAC From Registry

Write-Verbose "Disabling UAC In the Registry"
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -Value 0 | out-null
    
# Rename Computer

$Compname = read-host "Enter the new name of this Computer"
Rename-Computer -NewName "$CompName" -Restart


# Rename HDD's

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

