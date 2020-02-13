if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #


#License Expiration Dates Variables
function Test-LicenseVals{


  $LicenseKeyAExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyAExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyCExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyDExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyDExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyEExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyEExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyGExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyGExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyMExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyMExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyNExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyOExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyPExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPExpiry" -ErrorAction SilentlyContinue
  $LicenseKeyTExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTExpiry" -ErrorAction SilentlyContinue
  
  return $LicenseKeyAExpiry,$LicenseKeyCExpiry,$LicenseKeyDExpiry,$LicenseKeyEExpiry,$LicenseKeyGExpiry,$LicenseKeyMExpiry,$LicenseKeyNExpiry,$LicenseKeyOExpiry,$LicenseKeyPExpiry,$LicenseKeyTExpiry
  }
  
  Test-LicenseVals


$LicenseKeyDExpiry = Get-Item -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1"

If($null -eq $LicenseKeyDExpiry.GetValue("LicenseKeyDExpiry")) {
out-null
} 

Else {
$LicenseKeyDExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyDExpiry" -ErrorAction SilentlyContinue

If($LicenseKeyDExpiry) -ge (Get-Date -Format "MM/dd/yyyy")
{


}

}












Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting