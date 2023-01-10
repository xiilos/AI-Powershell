if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:

    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Registry Favorites
Start-Process Regedit

Write-Host "Creating Registry Favorites"
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Session Manager" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControl\Control\Session Manager" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "EnableLUA" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name ".Net Framework" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\.NETFramework" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "OpenDoor Software" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Add2Exchange"  -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Pendingfilerename" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "AutoDiscover" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Office" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Group Policy History" -Type string -Value "Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy\History" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Logon" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Windows Update" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Outlook Social Connector" -Type string -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect" -Force -ErrorAction SilentlyContinue
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit\Favorites" -Name "Modern Authentication" -Type string -Value "Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Force -ErrorAction SilentlyContinue
Start-Sleep -s 3
Stop-Process -Processname Regedit

Write-Host "Done"
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting