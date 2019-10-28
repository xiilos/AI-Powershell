if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Get-ADUser -Filter {Enabled -eq $false} -Properties showinaddressbook | Select-Object -ExpandProperty showinaddressbook


Get-Mailbox -Filter{(HiddenFromAddressListsEnabled -eq $true) -AND (UserAccountControl -eq "AccountDisabled, NormalAccount")}



Get-ADUser -Filter {(enabled -eq "false") -and (msExchHideFromAddressLists -notlike "*")} -searchbase "OU=users,DC=AD,DC=samenacapital,DC=com" -Properties enabled,msExchHideFromAddressLists




Get-ADUser `
 -Filter {(enabled -eq "false") -and (msExchHideFromAddressLists -notlike "*")} `
 -SearchBase "OU=<OrganisationalUnit>,DC=<Domain>,DC=<TLD>"`
 -Properties msExchHideFromAddressLists | `
 Set-ADUser -Add @{msExchHideFromAddressLists="TRUE"}



"OU=*,DC=AD,DC=samenacapital,DC=com"


get-aduser -filter {attribute -like 'Showinaddressbook'} -properties attribute

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting