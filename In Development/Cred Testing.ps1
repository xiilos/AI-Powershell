if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
                $ExchangeServerName = Read-Host "Please Type In your Exchange Server Name"
                $ExchangeUsername = Read-Host "Please Type In your Exchange Server Admin Username"
                $ExchangePassword = Read-Host "Please Type in your Exchange Server Admin Password" -AsSecureString
                $ExchangeServer = @{
                    Target   = "Exchange_Server"
                    UserName = "$ExchangeUsername"
                    Password = "$ExchangePassword"
                    Comment  = "$ExchangeServerName"
                    Persist  = "LocalMachine"
                }
                New-StoredCredential @ExchangeServer
            



# End Scripting