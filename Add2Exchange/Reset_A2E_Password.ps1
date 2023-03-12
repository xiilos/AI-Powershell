<#
        .SYNOPSIS
        Add2Exchange password Reset

        .DESCRIPTION
        Clears out the password field in A2E reg.
        Asks for new password and updates the Add2Exchange service with new password


        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
Write-Host "Reseting the Servive Account Password" -ForegroundColor Red
$Password = Read-Host "What is the New Service Account Password?"
$SVC = Get-WmiObject win32_service -Filter "Name='Add2Exchange Service'"
$SVC.StopService();
$Result = $SVC.Change($Null, $Null, $Null, $Null, $Null, $Null, $Null, "$Password")
If ($Result.ReturnValue -eq '0') { Write-Host "Add2Exchange Service Password Has Been Succsefully Changed" -ForegroundColor Green } Else { Write-Host "Error: $Result" }


Write-Host "Resetting the Add2Exchange Console Password" -ForegroundColor Red
Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "serviceAccountPwd" -Value "" -ErrorAction SilentlyContinue -ErrorVariable Error1
If ($Error1) {
    Write-Host "The Add2Exchange Console Password Could not be Changed" -ForegroundColor Red
}

Write-Host "Add2Exchange Console Password Has Been Succsefully Cleared" -ForegroundColor Green

Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting