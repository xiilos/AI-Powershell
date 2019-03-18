if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Group Policy Results
$TestPath = "C:\zlibrary\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Support"
}


do {
    Write-Host "Getting Group Policy Results"
    gpupdate
    gpresult /r >C:\zlibrary\Support\Group_policy_Report.txt
    Invoke-Item "C:\zlibrary\Support\Group_policy_Report.txt"

    Write-Host "If you get an error stating (No User Data in RSOP) you may have to edit the GPO History in REGEDIT"
    $confirmation = Read-Host "Would you like me to remove and update the GPO History for you? [Y/N]"
    if ($confirmation -eq 'y') {
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Group Policy\History' -Name DCName | out-null
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Group Policy\History' -Name PolicyOverdue | out-null
    }


    $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit