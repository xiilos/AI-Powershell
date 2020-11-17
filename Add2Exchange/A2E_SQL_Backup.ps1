if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Shutting Down Services
Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 2
Write-Host "Done"
Get-Service | Where-Object { $_.DisplayName -eq "Add2Exchange Service" } | Set-Service –StartupType Disabled

#Stop The Add2Exchange Agent
Write-Host "Stopping the Agent. Please Wait."
Start-Sleep -s 5
$Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
if ($Agent) {
    Write-Host "Waiting for Agent to Exit"
    Start-Sleep -s 5
    if (!$Agent.HasExited) {
        $Agent | Stop-Process -Force
    }
}
Write-Host "Stopping Add2Exchange SQL Service"
Stop-Service -Name "SQL Server (A2ESQLSERVER)"
Start-Sleep -s 2
Write-Host "Done"

#Backing Up SQL Files
Write-Host "Backinp Up Add2Exchange SQL Files"
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
Push-Location $Install
Copy-Item ".\Database\A2E.mdf" -Destination "C:\zlibrary\A2E_SQL_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB1
If ($DB1) {
    Write-Host "Error.....Cannot Find A2E.mdf Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
    Pause
}
Copy-Item ".\Database\A2E_log.LDF" -Destination "C:\zlibrary\A2E_Backup" -ErrorAction SilentlyContinue -ErrorVariable DB2
If ($DB2) {
    Write-Host "Error.....Cannot Find A2E_Log.LDF Database. Please Copy it Manually into the C:\zLibrary\A2E_Backup Folder"
    Pause
}






Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting