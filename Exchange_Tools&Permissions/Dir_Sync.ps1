if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #
$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Caution... You Must Run this on a box with Active Directory. If the box you are running this on does not have Active Directory; Click Cancel and the File will be Automatically copied to your Clipboard. Otherwise, Click OK to Continue.", 0, "WARNING!!", 0x1)
if ($answer -eq 2) {
    $Location = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
    Push-Location $Location
    Set-Clipboard -Path ".\Setup\Dir_Sync.ps1"
    Write-Host "File Copied"
    Pause
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}


$ADServer = Read-Host "Name of AD Server"
$session = New-PSSession -ComputerName $ADServer
Write-Host "Running Directory Sync"
Invoke-Command -Session $session -ScriptBlock {Import-Module -Name 'ADSync'}
Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
Remove-PSSession $session
Write-Host "Done"
Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting