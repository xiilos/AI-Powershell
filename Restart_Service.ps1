if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted

$SvcName = '3CXMediaServer'
$SvrName = '3CX-Phone'

#variables:
[string]$WaitForIt = ""
[string]$Verb = ""
[string]$Result = "FAILED"
$svc = (get-service -computername $SvrName -name $SvcName)
Write-host "$SvcName on $SvrName is $($svc.status)"
Switch ($svc.status) {
    'Stopped' {
        Write-host "Starting $SvcName..."
        $Verb = "start"
        $WaitForIt = 'Running'
        $svc.Start()}
    'Running' {
        Write-host "Restarting $SvcName..."
        $Verb = "Restart"
        $WaitForIt = 'Restarting'
        Restart-Service $SvcName -Force}
    Default {
        Write-host "$SvcName is $($svc.status).  Taking no action."}
}
if ($WaitForIt -ne "") {
    Try {  
        $svc.WaitForStatus($WaitForIt,'00:02:00')
    } Catch {
        Write-host "After waiting for 2 minutes, $SvcName failed to $Verb."
    }
    $svc = (get-service -computername $SvrName -name $SvcName)
    if ($svc.status -eq $WaitForIt) {$Result = 'SUCCESS'}
    Write-host "$Result`: $SvcName on $SvrName is $($svc.status)"
}

Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting