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
  Set-TransportService Exchange2019 -MessageTrackingLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\MessageTracking”
  Set-TransportService Exchange2019 -ConnectivityLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Connectivity”
  Set-TransportService Exchange2019 -IrmLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\IRMLogs”
  Set-TransportService Exchange2019 -ActiveUserStatisticsLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ActiveUserStats”
  Set-TransportService Exchange2019 -ServerStatisticsLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ServerStats”
  Set-TransportService Exchange2019 -ReceiveProtocolLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ProtocolLog\SmtpReceive”
  Set-TransportService Exchange2019 -RoutingTableLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Routing”
  Set-TransportService Exchange2019 -SendProtocolLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\ProtocolLog\SmtpSend”
  Set-TransportService Exchange2019 -TransportHttpLogPath “D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\TransportHttp”



  Set-TransportService Exchange2019 -QueueLogPath "D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\QueueViewer"
  Set-TransportService Exchange2019 -WlmLogPath "D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\WLM"
  Set-TransportService Exchange2019 -AgentLogPath "D:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\AgentLog"
  }

  #Check Log Location
  Get-TransportService exchange2019 | Select-Object *logpath

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting