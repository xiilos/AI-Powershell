if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

#Adding the WU Module in Powershell
Import-Module "PSWindowsUpdate" -ErrorAction SilentlyContinue
If ($error) {
    Write-Host "Adding Windows Update Module"
    Install-Module PSWindowsUpdate -Confirm:$False
} 
Else {
    Write-Host 'Module is installed'
    Update-Module "PSWindowsUpdate"
}


#Installing Windows Updates
Write-Host "Getting Windows Updates...Please Wait..."
Get-WindowsUpdate -Install -Confirm:$False –IgnoreReboot


#Cleaning
Write-Host "Cleaning Up....."
Write-Host "Flushing DNS"
IPConfig /FlushDNS
Write-Host "Done"

Write-Host "Cleaning up Files via Disk Cleanup"

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Active Setup Temp Folders" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\BranchCache" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\D3D Shader Cache"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\D3D Shader Cache" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Delivery Optimization Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Delivery Optimization Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Diagnostic Data Viewer database files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Diagnostic Data Viewer database files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Downloaded Program Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\DownloadsFolder"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\DownloadsFolder" /v StateFlags0100 /d 0 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Internet Cache Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Language Pack"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Language Pack" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Mixed Reality"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Mixed Reality" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Offline Pages Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Offline Pages Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Old ChkDsk Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Recycle Bin" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\RetailDemo Offline Content"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\RetailDemo Offline Content" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Service Pack Cleanup" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Setup Log Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error memory dump files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\System error minidump files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Setup Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Sync Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Sync Files" /V StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Thumbnail Cache" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Upgrade Discarded Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\User file versions"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\User file versions" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Defender"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Defender" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Error Reporting Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows ESD installation files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows ESD installation files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---

$TestPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files"
  
If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    REG ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files" /v StateFlags0100 /d 2 /t REG_DWORD /f
}
Else {
    Write-Host ""
}

#---


cleanmgr /sagerun:100 | Out-Null
Write-Host "Done"

#Update BGINFO
#Write-Host "Updating Info..."
#Start-Process "\\fileserv-db1\Sandbox\Port Apps\Documents\Quick Clean & Tools\BGinfo\standard.bgi" -ErrorAction SilentlyContinue

#Rebooting
Write-Host "Rebooting...."
Shutdown -r -t 10


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting