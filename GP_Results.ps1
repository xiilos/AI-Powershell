if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Group Policy Results

Write-Host "Getting Group Policy Results"
gpresult /r >c:\Group_policy_Report.txt
Invoke-Item "c:\Group_policy_Report.txt"

Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit