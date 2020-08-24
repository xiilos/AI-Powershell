if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

#Stopping Virtual Machines

#cmd.exe "C:\Program Files (x86)\VMware\VMware Workstation\vmrun" -T ws stop "C:\Virtual Machines\Mint\mint.vmx" soft

#cmd.exe "C:\Program Files (x86)\VMware\VMware Workstation\vmrun" -T ws stop "C:\Virtual Machines\Server 2019\Server 2019.vmx" soft

#cmd.exe "C:\Program Files (x86)\VMware\VMware Workstation\vmrun" -T ws stop "C:\Virtual Machines\SCO-Unix\SCO Unix IST.vmx" soft

pause
#Backing Up Virtual Machines

#Copy-Item "C:\Virtual Machines\Linux Mint" -Destination "C:\Backup\Linux Mint" -wait

#$Date = (get-date -format MMddyyyy)
#New-Item -Path "C:\Backup\Server 2019\Server2019$Date"
#Copy-Item "C:\Virtual Machines\Server 2019" -Recurse -Destination "C:\Backup\Server 2019\" -wait


Compress-Archive 'C:\Virtual Machines\Mint' -DestinationPath ('C:\Backup\Linux Mint\' + (get-date -Format MMddyyyy) + '.zip')

Compress-Archive 'C:\Virtual Machines\Server 2019' -DestinationPath ('C:\Backup\Server 2019\' + (get-date -Format MMddyyyy) + '.zip')

Compress-Archive 'C:\Virtual Machines\SCO-Unix' -DestinationPath ('C:\Backup\SCO-Unix\' + (get-date -Format MMddyyyy) + '.zip')

pause
#Starting Virtual Machines

cmd.exe "C:\Program Files (x86)\VMware\VMware Workstation\vmrun" -T ws start "C:\Virtual Machines\Linux Mint\mint.vmx"

cmd.exe "C:\Program Files (x86)\VMware\VMware Workstation\vmrun" -T ws start "C:\Virtual Machines\Server 2019\Windows Server 2019.vmx"

pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting