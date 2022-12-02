if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue


#Variables
$Nodes =  "Win10TestBox, A2E-O365, A2E-2022, Exchange2019, DC"

#Installing Windows Updates
Write-Host "Getting Windows Updates...Please Wait..."
Invoke-WUJob -ComputerName $Nodes -Script {Import-Module PSWindowsUpdate; Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot} -RunNow -Confirm:$false |  Out-File "\\file_server\Logs\MS Update Logs\$Nodes-$(Get-Date -f yyyy-MM-dd)-MSUpdates.log" -Force




Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting