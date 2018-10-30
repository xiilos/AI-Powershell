if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Unrestricted


Cd ùC:\Program Files (x86)\OpenDoor Softwareù\Add2Exchange\Setup Initializeù

C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "Windows Services.dll" /tlb:"Windows Services.tlb" /unregister
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "Windows Services.dll" /tlb:"Windows Services.tlb" /codebase



Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting