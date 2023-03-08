if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

# Define the network segment to scan
$networkSegment = "192.168.1.1"

# Loop through all IP addresses in the network segment
for ($i = 1; $i -le 255; $i++) {
    $ipAddress = "$networkSegment$i"

    # Try to connect to the remote computer
    if (Test-Connection -ComputerName $ipAddress -Count 1 -Quiet) {

        # Get basic information about the computer
        $computerName = (Get-WmiObject Win32_ComputerSystem -ComputerName $ipAddress).Name
        $osName = (Get-WmiObject Win32_OperatingSystem -ComputerName $ipAddress).Caption

        # Output the information to the console or a file
        Write-Output "$computerName, $ipAddress, $osName"
    }
}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting