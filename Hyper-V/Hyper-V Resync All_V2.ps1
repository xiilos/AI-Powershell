if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Script #

$Hosts = "Hyper-V08"
ForEach ($Server in $Hosts)
{
$Server
Invoke-Command -ComputerName $Server {
$FailedReplicas = Get-VMReplication | Where-Object{$_.Health -EQ 'Critical'}
ForEach ($VM in $FailedReplicas)
{
$VMName = $VM.Name
$VMName
Resume-VMReplication $VMName -Resynchronize
}
}
}

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting