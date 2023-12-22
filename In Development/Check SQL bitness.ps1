# Define the registry path for SQL Server instances
$instancesRegPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"

# Check if the path exists (for 64-bit installations on a 64-bit OS or any installation on a 32-bit OS)
if (Test-Path $instancesRegPath) {
    $instances = Get-ItemProperty $instancesRegPath
} else {
    # Check for 32-bit installations on a 64-bit OS
    $instancesRegPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL"
    if (Test-Path $instancesRegPath) {
        $instances = Get-ItemProperty $instancesRegPath
    }
}

# Check if any instances are found
if ($instances) {
    foreach ($instance in $instances.PSObject.Properties) {
        $instanceName = $instance.Name
        $sqlInstanceRegPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\" + $instance.Value + "\Setup"

        # Check for 32-bit instances on a 64-bit OS
        if (-not (Test-Path $sqlInstanceRegPath)) {
            $sqlInstanceRegPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\" + $instance.Value + "\Setup"
        }

        if (Test-Path $sqlInstanceRegPath) {
            $sqlInstance = Get-ItemProperty $sqlInstanceRegPath
            $version = $sqlInstance.Version
            $edition = $sqlInstance.Edition

            if ($edition -like "*64-bit*") {
                $bitness = "64-bit"
            } elseif ($edition -like "*32-bit*") {
                $bitness = "32-bit"
            } else {
                $bitness = "Unknown bitness"
            }

            Write-Host "Instance Name: $instanceName"
            Write-Host "Version: $version"
            Write-Host "Edition: $edition"
            Write-Host "Bitness: $bitness"
            Write-Host ""
            Pause
        }
    }
} else {
    Write-Host "No SQL Server instances found."
}
