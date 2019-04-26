if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Group Policy Results
$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Support"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Support Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Support"
}


Do {
    Write-Host "Getting Group Policy Results"
    gpupdate
    Start-Sleep -s 3
    gpresult /Scope Computer /v | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Support\Group_policy_Report.txt" -ErrorAction SilentlyContinue -ErrorVariable GPError;
    If ($GPError) { 
        Write-Warning -Message "(No User Data in RSOP) Updating GPO History...."
        Start-Sleep -s 3
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Group Policy\History' -Name DCName | out-null
        Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Group Policy\History' -Name PolicyOverdue | out-null
        gpresult /Scope Computer /v | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Support\Group_policy_Report.txt" -ErrorAction SilentlyContinue
    }

    Start-Sleep -s 3
    Invoke-Item "C:\Program Files (x86)\DidItBetterSoftware\Support\Group_policy_Report.txt"

    $repeat = Read-Host 'Do you want to run it again? [Y/N]'

} Until ($repeat -eq 'n')

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit