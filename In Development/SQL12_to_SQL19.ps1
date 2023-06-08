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

# Define paths and file names
$setupFilePath = "C:\Path\To\SQL2019\setup.exe"
$configurationFilePath = "C:\Path\To\Configuration.ini"
$logFilePath = "C:\Path\To\UpgradeLog.txt"

# Define SQL Server instance names
$instanceName = "SQLEXPRESS"

# Define the product and feature codes
$productCode = "{0A869A65-8C94-4F7C-A5C7-972D3C8CED9E}"
$featureCode = "SQL"

# Define the upgrade action
$upgradeAction = "Upgrade"

# Define the log level
$logLevel = "Standard"

# Run the SQL Server setup
Start-Process -FilePath $setupFilePath -ArgumentList "/Q /IACCEPTSQLSERVERLICENSETERMS /ConfigurationFile=`"$configurationFilePath`" /INDICATEPROGRESS /ACTION=`"$upgradeAction`" /INSTANCEID=`"$instanceName`" /FEATURES=`"$featureCode`" /INSTANCENAME=`"$instanceName`" /SQLSYSADMINACCOUNTS=`"BUILTIN\Administrators`" /SAPWD=`"NewPassword`" /SQLCOLLATION=`"SQL_Latin1_General_CP1_CI_AS`" /ERRORREPORTING=`"False`" /QUIET /IAcceptSQLServerLicenseTerms" -Wait

# Check the exit code to determine if the upgrade was successful
$exitCode = $LASTEXITCODE
if ($exitCode -eq 0) {
    Write-Host "SQL Server upgrade completed successfully."
} else {
    Write-Host "SQL Server upgrade failed. Exit code: $exitCode"
}

# Write the setup log to a file
$setupLog = Get-Content -Path "$env:TEMP\SqlSetup.log" -Raw
$setupLog | Out-File -FilePath $logFilePath -Encoding UTF8
Write-Host "Setup log saved to: $logFilePath"


Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting




#Alternate

# Variables
$sourceInstance = "SQLEXPRESS"
$targetInstance = "SQLEXPRESS"
$sourceVersion = "2012"
$targetVersion = "2019"
$installPath = "C:\SQL2019"

# Download SQL Server 2019 Express
$downloadUrl = "https://go.microsoft.com/fwlink/?linkid=866658"
$installerPath = "$installPath\SQLEXPR.exe"

Write-Host "Downloading SQL Server 2019 Express..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

# Install SQL Server 2019 Express
Write-Host "Installing SQL Server 2019 Express..."
Start-Process -FilePath $installerPath -ArgumentList "/ACTION=Upgrade /INSTANCENAME=$sourceInstance /Q /IACCEPTSQLSERVERLICENSETERMS" -Wait

# Verify the installation
$targetInstancePath = "C:\Program Files\Microsoft SQL Server\MSSQL$targetInstance"

if (Test-Path $targetInstancePath) {
    Write-Host "SQL Server $targetVersion Express installation completed successfully."
} else {
    Write-Host "SQL Server $targetVersion Express installation failed."
}
