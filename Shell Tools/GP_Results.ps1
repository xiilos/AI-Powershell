if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

# Group Policy Results
$TestPath = "C:\zlibrary"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "zLibrary Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary"
}


Do {
    Write-Host "Getting Group Policy Results"
    gpupdate
    Start-Sleep -s 2
    gpresult /Scope Computer /h C:\zlibrary\Computer_Group_policy_Report.html /f
    gpresult /Scope User /h C:\zlibrary\User_Group_policy_Report.html /f

    Invoke-Item "C:\zlibrary\Computer_Group_policy_Report.html"
    Invoke-Item "C:\zlibrary\User_Group_policy_Report.html"

    $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit