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

# Load the SQL Server module
Import-Module SqlServer

# Define your SQL Server instance and database name
$instanceName = "A2ESQLSERVER"
$databaseName = "A2E"

# The SQL command to shrink the log file
$sqlCommand = @"
USE [$databaseName]
GO
DBCC SHRINKFILE (N'$databaseName_log' , 0, TRUNCATEONLY)
GO
"@

# Execute the command
Invoke-Sqlcmd -Query $sqlCommand -ServerInstance $instanceName

Write-Host "Shrink operation completed."






# PowerShell script to trim .ldf file

# SQL Server details
$serverName = "A2ESQLSERVER"
$databaseName = "A2E"
$ldfFileName = "A2E_LOG"


# Load the required assembly
Add-Type -AssemblyName "Microsoft.SqlServer.Smo"
$server = New-Object Microsoft.SqlServer.Management.Smo.Server $serverName


# Shrink the .ldf file
$query = "DBCC SHRINKFILE('$ldfFileName', 1)" # 1 MB is used as an example minimum size
$server.Databases[$databaseName].ExecuteNonQuery($query)

Write-Output "Transaction log has been backed up and shrunk!"







# Load the required SMO assembly
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null

# Set your SQL Server instance and database details
$serverName = "localhost" # Example: "localhost" or "localhost\SQLExpress"
$databaseName = "A2E"
$credentials = "Integrated Security=SSPI" # For Windows Authentication. Otherwise, you might need to provide username and password.

# Create a server object
$server = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $serverName
$server.ConnectionContext.LoginSecure=$true
$server.ConnectionContext.ConnectionString = "Server=$serverName; $credentials"

# Check if the database exists
if ($server.Databases[$databaseName]) {
    # Detach the database
    $server.DetachDatabase($databaseName, $false)
    Write-Output "Database $databaseName detached successfully."
} else {
    Write-Output "Database $databaseName does not exist."
}





Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting