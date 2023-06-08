<#
        .SYNOPSIS
        

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

# Script #

# Path to the SQL Server 2019 Express installation media
$installationMediaPath = "C:\path\to\SQL2019Express.iso"

# Mount the SQL Server 2019 Express installation media as a virtual drive
Mount-DiskImage -ImagePath $installationMediaPath

# Specify the path to the setup.exe file within the mounted ISO
$setupPath = (Get-DiskImage -ImagePath $installationMediaPath | Get-Volume).DriveLetter + ":\setup.exe"

# Specify the path for the SQL Server 2019 installation log file
$installLogFile = "C:\Temp\SQL2019_Upgrade_Log.txt"

# Set the SQL Server instance name
$instanceName = "MSSQLSERVER"

# Specify the SQL Server 2019 product key (optional)
$productKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"

# Specify the SQL Server 2019 edition (optional)
# Available editions: "Express", "Standard", "Enterprise"
$edition = "Express"

# Specify the SQL Server 2019 installation configuration file (optional)
$configurationFile = "C:\path\to\ConfigurationFile.ini"

# Set the SQL Server 2017 instance name
$sourceInstanceName = "MSSQLSERVER"

# Specify the path for the SQL Server 2017 installation log file
$sourceInstallLogFile = "C:\Temp\SQL2017_Upgrade_Log.txt"

# Start the SQL Server 2019 Express installation
Start-Process -FilePath $setupPath -ArgumentList "/QUIET", "/ACTION=Upgrade", "/INSTANCENAME=$instanceName", "/IACCEPTSQLSERVERLICENSETERMS", "/ENU", "/INDICATEPROGRESS", "/ERRORREPORTING=1", "/SQLSYSADMINACCOUNTS=BUILTIN\Administrators", "/SAPWD=YourStrongPassword", "/ConfigurationFile=$configurationFile", "/UIMODE=Normal" -Wait -NoNewWindow

# Unmount the SQL Server 2019 Express installation media
Dismount-DiskImage -ImagePath $installationMediaPath

# Restart the SQL Server 2019 service
Restart-Service -Name "MSSQL`$$instanceName" -Force

# Check the SQL Server 2019 service status
$serviceStatus = Get-Service -Name "MSSQL`$$instanceName"

if ($serviceStatus.Status -eq "Running") {
    Write-Host "SQL Server 2019 upgrade completed successfully."
} else {
    Write-Host "SQL Server 2019 upgrade failed. Check the installation log file: $installLogFile"
}


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting