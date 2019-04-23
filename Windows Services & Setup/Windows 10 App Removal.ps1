if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
Write-Host "Removing Windows Apps"
Get-AppxPackage -AllUsers | where-object {$_.name -notlike "*windows.photos"} | where-object {$_.name -notlike "*store*"} | where-object {$_.name -notlike "*calculator*"} | where-object {$_.name -notlike "*sticky*"} | where-object {$_.name -notlike "*soundrecorder*"} | where-object {$_.name -notlike "*mspaint*"} | where-object {$_.name -notlike "*screensketch*"} | Remove-AppxPackage -Confirm:$False -ErrorAction SilentlyContinue -ErrorVariable ProcessError;

If ($ProcessError) {

    write-warning -message "Cannot Remove some Apps. Maybe not a windows App?";
}


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting