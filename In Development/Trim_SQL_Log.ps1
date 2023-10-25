<#
        .SYNOPSIS
        Trims SQL DB A2E_Log.LDF

        .DESCRIPTION
      

        .NOTES
        Version:        1.0
        Author:         DidItBetter Software

    #>


if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

#Variables
$ServerName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "Server" -ErrorAction SilentlyContinue
$instanceName = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "DBInstance" -ErrorAction SilentlyContinue
$DBname = "A2E"
#$Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
#$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring



# Script #
Write-Host "Trimming the SQL transaction Log"

try {
    Invoke-Sqlcmd -ServerInstance "$Servername\$instancename" -Database "$DBname" -Query "DBCC SHRINKFILE('A2E_log', 1);"
}
catch {
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10020 -EntryType FailureAudit -Message "SQL Transaction Log Trim failure $_.Exception.Message"
    Write-Host "Please see the Add2Exchange Log for Errors"
    Pause
    Get-PSSession | Remove-PSSession
    Exit
}

Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10021 -EntryType FailureAudit -Message "Add2Exchange SQL Transaction Log Trimmed Succesfully"


#Invoke-Sqlcmd -ServerInstance "$Servername\$instancename" -Database "$DBname" -Username "$username" -Password "$password" or -Trustservercertificate -Query "DBCC SHRINKFILE('A2E_log', 1);"

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting

