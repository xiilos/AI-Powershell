if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Script #

#Auto Reboot
$Confirmation = Read-Host "This update may need a reboot. Reboot automatically after successful Install? [Y/N]"
If ($confirmation -eq 'y') { $Reboot = "Auto Reboot Selected" }
If ($confirmation -eq 'n') { $NoReboot = "Please reboot when possible after update" }


#Create zLibrary
Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\.NET Updates"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\.NET Updates"
}

#Check Exchange Online Management Module for updates
Write-Host "Checking for Exhange Online Module"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

IF (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "Exchange Online Module Exists"

    $InstalledEXOv2 = ((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()

    $LatestEXOv2 = (Find-Module -Name ExchangeOnlineManagement).Version.ToString()

    [PSCustomObject]@{
        Match = If ($InstalledEXOv2 -eq $LatestEXOv2) { Write-Host "You are on the latest Version" } 

        Else {
            Write-Host "Upgrading Modules..."
            Update-Module -Name ExchangeOnlineManagement -Force
            Write-Host "Success"
        }

    }


} 
Else {
    Write-Host "Module Does Not Exist"
    Write-Host "Downloading Exchange Online Management..."
    Install-Module –Name ExchangeOnlineManagement -Force
    Write-Host "Success"
}



# .NET 4.8 or higher
If ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version -ge '4.8') {
    $version = ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version)
    Write-Host ".NET $version is installed and meets Add2Exchange Requirements" -ForegroundColor Green
    Pause
}
Else {
    $version = ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version)
    Write-Host ".NET $version is installed and needs to be upgraded" -ForegroundColor Red
    Write-Host "Updating .NET Framework"

    #Shutting Down Services
    Write-Host "Stopping Add2Exchange Service"
    Stop-Service -Name "Add2Exchange Service"
    Start-Sleep -s 10
    Write-Host "Done"
    

    #Stop The Add2Exchange Agent
    Write-Host "Stopping the Agent. Please Wait."
    Start-Sleep -s 10
    $Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
    if ($Agent) {
        Write-Host "Waiting for Agent to Exit"
        Start-Sleep -s 10
        if (!$Agent.HasExited) {
            $Agent | Stop-Process -Force
        }
    }
    Write-Host "Stopping Add2Exchange SQL Service"
    Stop-Service -Name "SQL Server (A2ESQLSERVER)"
    Start-Sleep -s 10
    Write-Host "Done"


    #Downloading .NET
    Write-Host "Downloading .NET"
    $URL = "https://go.microsoft.com/fwlink/?linkid=2088631"
    $Output = "c:\zlibrary\.NET Updates\ndp48-x86-x64-allos-enu.exe"

        (New-Object System.Net.WebClient).DownloadFile($URL, $Output)


    #Installing .NET
    Write-Host "Installing .NET..Please Wait..."
    If ($Reboot) {
        Start-Process "C:\zlibrary\.NET Updates\ndp48-x86-x64-allos-enu.exe" -ArgumentList "Setup /q /norestart" -Wait
        Write-Host "Update complete. Rebooting in 10 Seconds"
        Shutdown -r -t 10
    }

    If ($NoReboot) {
        Start-Process "C:\zlibrary\.NET Updates\ndp48-x86-x64-allos-enu.exe" -ArgumentList "Setup /q /norestart" -Wait
        Write-Host "Update complete. Please Reboot when possible."
        Pause
    }
    

}
    

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting