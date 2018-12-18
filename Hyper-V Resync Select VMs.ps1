if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Notes: This script resumes all replication
# Script #

Resume-VMReplication –VMName "3cx-phone" -Resynchronize -wait
Resume-VMReplication –VMName "A2E" -Resynchronize -wait
Resume-VMReplication –VMName "A2E MaidenbaumandSternberg 8367" -Resynchronize -wait
Resume-VMReplication –VMName "A2O MHR Funds 2914" -Resynchronize -wait
Resume-VMReplication –VMName "A2O PCAM 9705" -Resynchronize -wait
Resume-VMReplication –VMName "A2O Sarkela 1197" -Resynchronize -wait
Resume-VMReplication –VMName "Colorado-Legislative-2089" -Resynchronize -wait
Resume-VMReplication –VMName "cPanel" -Resynchronize -wait
Resume-VMReplication –VMName "DIBDC" -Resynchronize -wait
Resume-VMReplication –VMName "DIBDC1" -Resynchronize -wait
Resume-VMReplication –VMName "DIBEX10" -Resynchronize -wait
Resume-VMReplication –VMName "Docker-a" -Resynchronize -wait
Resume-VMReplication –VMName "eset" -Resynchronize -wait
Resume-VMReplication –VMName "FTP / WSUS" -Resynchronize -wait
Resume-VMReplication –VMName "H4101 Very Rare A2E" -Resynchronize -wait
Resume-VMReplication –VMName "H4102 2562 Sutton A2E" -Resynchronize -wait
Resume-VMReplication –VMName "H4103 9461 Wimmer Bros A2E" -Resynchronize -wait
Resume-VMReplication –VMName "H4104" -Resynchronize -wait
Resume-VMReplication –VMName "H4105 7329 Steadfast A2E" -Resynchronize -wait
Resume-VMReplication –VMName "H4106 9096 Balance Tech A2E" -Resynchronize -wait
Resume-VMReplication –VMName "H4107" -Resynchronize -wait
Resume-VMReplication –VMName "HOSTED EX2016" -Resynchronize -wait
Resume-VMReplication –VMName "IIS" -Resynchronize -wait
Resume-VMReplication –VMName "MOJO-B" -Resynchronize -wait
Resume-VMReplication –VMName "Monitor" -Resynchronize -wait
Resume-VMReplication –VMName "Redmine" -Resynchronize -wait
Resume-VMReplication –VMName "Screenconnect" -Resynchronize -wait
Resume-VMReplication –VMName "Shared DC - 2016" -Resynchronize -wait
Resume-VMReplication –VMName "SpamTitan" -Resynchronize -wait
Resume-VMReplication –VMName "Spree-B" -Resynchronize -wait
Resume-VMReplication –VMName "SQL" -Resynchronize -wait
Resume-VMReplication –VMName "SuperOutlook" -Resynchronize -wait
Resume-VMReplication –VMName "Syslog" -Resynchronize -wait
Resume-VMReplication –VMName "Veeam" -Resynchronize -wait
Resume-VMReplication –VMName "W12-A2E-FitnessFrame" -Resynchronize -wait


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting