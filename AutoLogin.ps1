if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


$message  = 'Please Pick How to Logon'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Enable Auto Login'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Disable Auto Login'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

if ($decision -eq 2) {
	Write-Output "Quitting"
	Get-PSSession | Remove-PSSession
	Exit
}



# Enable Auto Login

if ($decision -eq 0) {

do {
$UserName = read-host "Enter the Username Example: DOMAIN\USERNAME";
$Password = read-host "Enter the Account password"



	Write-Verbose "Enabling Windows AutoLogon"
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 | out-null
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value $UserName | out-null
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value $Password | out-null

  $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}




# Disable Auto Login


if ($decision -eq 1) {


	Write-Verbose "Disabling Windows AutoLogon"
	Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon | out-null
	Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName | out-null
	Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword | out-null

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit
}
