if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

$Enabled = Get-MpComputerStatus | Select-Object -expandproperty realtimeprotectionenabled

if ($Enabled -eq $True) {

  Write-Host "Gathering Information about your Installation..."

  $InstallLocation = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue

  $Drive = (get-item $InstallLocation).parent.parent.parent | Select-Object name -expandproperty name

  #Variables:
  $EXCL1 = join-path -path $Drive -childpath "Program Files (x86)\OpenDoor Software®"
  $EXCL2 = join-path -path $Drive -childpath "Program Files (x86)\Microsoft SQL Server"
  $EXCL3 = join-path -path $Drive -childpath "Program Files\Microsoft SQL Server"
  $EXCL4 = join-path -path $Drive -childpath "Program Files (x86)\DidItBetterSoftware"
  $EXCL5 = join-path -path $Drive -childpath "zLibrary"
  $EXCL6 = join-path -path $Drive -childpath "Program Files (x86)\Microsoft Office"
  $EXCL7 = "C:\Users\zadd2exchange\AppData"



  #Write the Exclusions
  Write-Host "Creating Exclusions for Windows Defender"
  Add-MpPreference -ExclusionPath $EXCL1
  Add-MpPreference -ExclusionPath $EXCL2
  Add-MpPreference -ExclusionPath $EXCL3
  Add-MpPreference -ExclusionPath $EXCL4
  Add-MpPreference -ExclusionPath $EXCL5
  Add-MpPreference -ExclusionPath $EXCL6
  Add-MpPreference -ExclusionPath $EXCL7

  Write-Host "Done"

  Write-Host "ttyl"
  Get-PSSession | Remove-PSSession
  Exit


}
Else {
  $wshell = New-Object -ComObject Wscript.Shell
  $wshell.Popup("Windows Defender is Disabled (No Changes Made)", 0, "Windows Defender", 0x1)
}

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit



# End Scripting