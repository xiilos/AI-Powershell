if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Notes: Resumes and Resyncs all Missions Critical VMs

# Script #

Resume-VMReplication –VMName "3cx-phone" -Resynchronize
Resume-VMReplication –VMName "DIBDC" -Resynchronize
Resume-VMReplication –VMName "DIBDC1" -Resynchronize
Resume-VMReplication –VMName "DIBEX10" -Resynchronize
Resume-VMReplication –VMName "Docker-a" -Resynchronize
Resume-VMReplication –VMName "HOSTED EX2016" -Resynchronize
Resume-VMReplication –VMName "IIS" -Resynchronize
Resume-VMReplication –VMName "MOJO-B" -Resynchronize
Resume-VMReplication –VMName "Redmine" -Resynchronize
Resume-VMReplication –VMName "Shared DC - 2016" -Resynchronize
Resume-VMReplication –VMName "SpamTitan" -Resynchronize
Resume-VMReplication –VMName "Spree-B" -Resynchronize
Resume-VMReplication –VMName "SQL" -Resynchronize
Resume-VMReplication –VMName "Veeam" -Resynchronize


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting