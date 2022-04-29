

$Repeater = (New-TimeSpan -Minutes 360)
$Duration = ([timeSpan]::maxvalue)
$Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to All O365" -Description "Add Permissions to All Users in Office 365"


#Add Below to all tasks
$UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to All O365" -Description "Add Permissions to All Users in Office 365" -User $UserID -Password $Password



$Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
Set-Location $Location

$Repeater = (New-TimeSpan -Minutes 360)
$Duration = ([timeSpan]::maxvalue)
$Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
$UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
$Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to All O365" -Description "Add Permissions to All Users in Office 365" -User $UserID -Password $Password