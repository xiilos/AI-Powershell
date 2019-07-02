if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Script #

Stop-Computer -ComputerName "Hyper-V14"
Write-Host "Waiting for Hyper-V14 to go Down"
Start-Sleep -s 60
If (Test-Connection -ComputerName "Hyper-V14" -Quiet) {
    Write-Host "Hyper-V14 is Still Up!"
}

Else {
    Write-Host "Waiting for Hyper-V14 to go Down"
    Start-Sleep -s 60
    Test-Connection -ComputerName "Hyper-V14"
}
#------------------------------------------------------------------------------
Stop-Computer -ComputerName "Hyper-V13"
Write-Host "Waiting for Hyper-V13 to go Down"
Start-Sleep -s 60
If (Test-Connection -ComputerName "Hyper-V13" -Quiet) {
    Write-Host "Hyper-V13 is Still Up!"
}

Else {
    Write-Host "Waiting for Hyper-V13 to go Down"
    Start-Sleep -s 60
    Test-Connection -ComputerName "Hyper-V13"
}
   
#------------------------------------------------------------------------------
Stop-Computer -ComputerName "Hyper-V12"
Write-Host "Waiting for Hyper-V12 to go Down"
Start-Sleep -s 60
If (Test-Connection -ComputerName "Hyper-V12" -Quiet) {
    Write-Host "Hyper-V12 is Still Up!"
}

Else {
    Write-Host "Waiting for Hyper-V12 to go Down"
    Start-Sleep -s 60
    Test-Connection -ComputerName "Hyper-V12"
}

#------------------------------------------------------------------------------
Stop-Computer -ComputerName "Hyper-V11"
Write-Host "Waiting for Hyper-V11 to go Down"
Start-Sleep -s 60
If (Test-Connection -ComputerName "Hyper-V11" -Quiet) {
    Write-Host "Hyper-V11 is Still Up!"
}

Else {
    Write-Host "Waiting for Hyper-V11 to go Down"
    Start-Sleep -s 60
    Test-Connection -ComputerName "Hyper-V11"
}


#------------------------------------------------------------------------------
Stop-Computer -ComputerName "Hyper-V10"
Write-Host "Waiting for Hyper-V10 to go Down"
Start-Sleep -s 60
If (Test-Connection -ComputerName "Hyper-V10" -Quiet) {
    Write-Host "Hyper-V10 is Still Up!"
}

Else {
    Write-Host "Waiting for Hyper-V10 to go Down"
    Start-Sleep -s 60
    Test-Connection -ComputerName "Hyper-V10"
}

Write-Host "Done"
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting