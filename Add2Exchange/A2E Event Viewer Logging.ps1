if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

$LogName = "Add2Exchange"
$EventLogName = "A2EServiceTracker"
$EventLogIDFail = 9001
$EventLogIDSuccess = 9000
$NumHours = 1
$RegData = Get-ItemProperty -Path HKLM:\Software\A2EServiceMonitor
$LastDate = $RegData.LastSync
$CurrentDate = Get-Date
$Difference = New-TimeSpan -Start $LastDate -End $CurrentDate

If (![System.Diagnostics.EventLog]::SourceExists($EventLogName))
{
       New-EventLog -LogName $LogName -Source $EventLogName
}

If ($Difference.Hours -ge $NumHours)
{
       Write-EventLog -LogName $LogName -Source $EventLogName -EventID $EventLogIDFail -EntryType Error -Category 0 -Message "Add2Exchange has not run since $LastDate... Killing Agent Process"
       Stop-Process -Name "Add2Exchange Agent" -Force
}
else {
       Write-EventLog -LogName $LogName -Source $EventLogName -EventID $EventLogIDSuccess -EntryType Information -Category 0 -Message "Add2Exchange synced recently at $LastDate"
}



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting