if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Notes: This script resumes all replication
# Script #

Resume-VMReplication –VMName "3cx-phone" -Resynchronize
Resume-VMReplication –VMName "A2E" -Resynchronize
Resume-VMReplication –VMName "A2E MaidenbaumandSternberg 8367" -Resynchronize
Resume-VMReplication –VMName "A2O MHR Funds 2914" -Resynchronize
Resume-VMReplication –VMName "A2O PCAM 9705" -Resynchronize
Resume-VMReplication –VMName "A2O Sarkela 1197" -Resynchronize
Resume-VMReplication –VMName "Colorado-Legislative-2089" -Resynchronize
Resume-VMReplication –VMName "cPanel" -Resynchronize
Resume-VMReplication –VMName "DIBDC" -Resynchronize
Resume-VMReplication –VMName "DIBDC1" -Resynchronize
Resume-VMReplication –VMName "DIBEX10" -Resynchronize
Resume-VMReplication –VMName "Docker-a" -Resynchronize
Resume-VMReplication –VMName "eset" -Resynchronize
Resume-VMReplication –VMName "FTP / WSUS" -Resynchronize
Resume-VMReplication –VMName "H4101 Very Rare A2E" -Resynchronize
Resume-VMReplication –VMName "H4102 2562 Sutton A2E" -Resynchronize
Resume-VMReplication –VMName "H4103 9461 Wimmer Bros A2E" -Resynchronize
Resume-VMReplication –VMName "H4104" -Resynchronize
Resume-VMReplication –VMName "H4105 7329 Steadfast A2E" -Resynchronize
Resume-VMReplication –VMName "H4106 9096 Balance Tech A2E" -Resynchronize
Resume-VMReplication –VMName "H4107" -Resynchronize
Resume-VMReplication –VMName "HOSTED EX2016" -Resynchronize
Resume-VMReplication –VMName "IIS" -Resynchronize
Resume-VMReplication –VMName "MOJO-B" -Resynchronize
Resume-VMReplication –VMName "Monitor" -Resynchronize
Resume-VMReplication –VMName "Redmine" -Resynchronize
Resume-VMReplication –VMName "Screenconnect" -Resynchronize
Resume-VMReplication –VMName "Shared DC - 2016" -Resynchronize
Resume-VMReplication –VMName "SpamTitan" -Resynchronize
Resume-VMReplication –VMName "Spree-B" -Resynchronize
Resume-VMReplication –VMName "SQL" -Resynchronize
Resume-VMReplication –VMName "SuperOutlook" -Resynchronize
Resume-VMReplication –VMName "Syslog" -Resynchronize
Resume-VMReplication –VMName "Veeam" -Resynchronize
Resume-VMReplication –VMName "W12-A2E-FitnessFrame" -Resynchronize


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting