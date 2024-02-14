# Import the SQL Server module
Import-Module SqlServer

# Define the SQL Server instance and the database to detach
$serverInstance = "A2ESQLSERVER"
$databaseName = "A2E"

# Create a Server object
$server = New-Object Microsoft.SqlServer.Management.Smo.Server($serverInstance)

# Access the database
$database = $server.Databases[$databaseName]

if ($null -eq $database) {
    Write-Host "Database not found."
    exit
}

# Detach the database
$server.DetachDatabase($databaseName, $true, $false)

Write-Host "Database detached successfully."


#Re attach

# Import the SQL Server module
Import-Module SqlServer

# Define the SQL Server instance and the database details
$serverInstance = "A2ESQLSERVER"
$databaseName = "A2E"
$mdfFilePath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Database\A2E.MDF"
$ldfFilePath = "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Database\A2E_Log.LDF"

# Create a Server object
$server = New-Object Microsoft.SqlServer.Management.Smo.Server($serverInstance)

# Define the filegroups for the database
$dbFile = New-Object Microsoft.SqlServer.Management.Smo.DataFile($server.Databases[$databaseName].FileGroups[0], $databaseName)
$dbFile.FileName = $mdfFilePath

$logFile = New-Object Microsoft.SqlServer.Management.Smo.LogFile($server.Databases[$databaseName].LogFiles[0], $databaseName)
$logFile.FileName = $ldfFilePath

# Attach the database
$server.AttachDatabase($databaseName, [string[]]($mdfFilePath, $ldfFilePath))

Write-Host "Database attached successfully."

pause
