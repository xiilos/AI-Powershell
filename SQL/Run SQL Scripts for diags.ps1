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

$serverName = "A2ESQLSERVER"
$databaseName = "your_database_name"
$sqlScriptPath = "C:\path\to\your\script.sql"

# Create a connection to the SQL Server instance
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True"

# Open the connection
$sqlConnection.Open()

# Read the contents of the SQL script file
$sqlScript = Get-Content $sqlScriptPath -Raw

# Create a SQL command object and set its properties
$sqlCommand = New-Object System.Data.SqlClient.SqlCommand
$sqlCommand.Connection = $sqlConnection
$sqlCommand.CommandType = [System.Data.CommandType]::Text
$sqlCommand.CommandText = $sqlScript

# Execute the SQL script
$sqlCommand.ExecuteNonQuery()

# Close the connection
$sqlConnection.Close()


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting