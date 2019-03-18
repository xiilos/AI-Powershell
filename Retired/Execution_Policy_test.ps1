if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy
# Script #

$Policy = "RemoteSigned"
If ((get-ExecutionPolicy) -ne $Policy) {
   Set-ExecutionPolicy -Scope localmachine -ExecutionPolicy $Policy -Force
   Write-Host "Execution Policy is now Set to RemoteSigned"
  Exit
}



$Policy = "RemoteSigned"

if ( $(Try { Get-ExecutionPolicy $Policy.trim() } Catch { $false }) ) {
Write-Host "Execution Policy is already set to RemoteSigned"
 }
Else {
Set-ExecutionPolicy -Scope localmachine -ExecutionPolicy $Policy -Force
Write-Host "Execution Policy is now Set to RemoteSigned"
 }


# Script #





# End Scripting