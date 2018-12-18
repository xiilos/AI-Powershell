if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


# Script #
$Exchangename = Read-Host "What is your Exchange server name? (FQDN)"
Test-WSMan -ComputerName $Exchangename


        if ((test-wsman -ComputerName $Exchangename) -eq "1")
        {
        return "PSRemoting is enabled"
        }

            else
            {
            return $wshell = New-Object -ComObject Wscript.Shell

            $wshell.Popup("Before Continuing, please remote into your Exchange server.
            Open Powershell as administrator
            Type: *Enable-PSRemoting* without the stars and hit enter.
            Once Done, click OK to Continue",0,"Enable PSRemoting",0x1)
            }

# End Scripting