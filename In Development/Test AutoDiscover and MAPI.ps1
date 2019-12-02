if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Test-OutlookWebServices -identity zadd2exchange –MailboxCredential (Get-Credential)

Get-AutodiscoverVirtualDirectory -Identity "exchange2016\Autodiscover*" | Format-List

Get-AutodiscoverVirtualDirectory -Server cs-srv1

Get-AutodiscoverVirtualDirectory


#reset for public or 
#network location awarness


Test-MapiConnectivity -Server "Server01"

Test-MapiConnectivity -Identity "midwest\john"

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting