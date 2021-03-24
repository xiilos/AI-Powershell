if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

Get-TransportService | ForEach-Object {
  Set-TransportService Exchange2016 -MessageTrackingLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\MessageTracking”
  Set-TransportService Exchange2016 -ConnectivityLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Connectivity”
  Set-TransportService Exchange2016 -IrmLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\IRMLogs”
  Set-TransportService Exchange2016 -ActiveUserStatisticsLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ActiveUserStats”
  Set-TransportService Exchange2016 -ServerStatisticsLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ServerStats”
  Set-TransportService Exchange2016 -ReceiveProtocolLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ProtocolLog\SmtpReceive”
  Set-TransportService Exchange2016 -RoutingTableLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Routing”
  Set-TransportService Exchange2016 -SendProtocolLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ProtocolLog\SmtpSend”
  }

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting