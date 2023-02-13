if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

#Checking Task Creation

Write-Host "Checking for Auto update Task..."
$TaskName = "Scheduled Update Add2Exchange"
$TaskExists = Get-ScheduledTask | Where-Object { $_.TaskName -like $TaskName }

If ($TaskExists) {
    Write-Host "Task Exists" 
} 

Else {
    $Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
    Set-Location $Location
    #$Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1)
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Scheduled_Update_Add2Exchange.ps1"'
    $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
    Register-ScheduledTask -Action $Action -RunLevel Highest -TaskName "Scheduled Update Add2Exchange" -Description "Updates Add2Exchange to the latest Version" -User $UserID -Password $Password
    Write-Host "Done" 
}

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting