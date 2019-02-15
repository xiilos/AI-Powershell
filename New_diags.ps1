if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

$message  = 'SQL Tools'
$question = 'Pick one of the following from below'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&0 Check for Bad Rows'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&1 Remove Bad Rows'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&2 Exit'))



$decision = $Host.UI.PromptForChoice($message, $question, $choices, 2)

# Option 0: Check for Bad Rows

if ($decision -eq 0) {
    
    Set-Location "SQLSERVER:\SQL\MyComputer\MainInstance"
    Invoke-Sqlcmd -Query "SELECT * FROM A2E_UserProperties AS LFT
    WHERE A2EDeleted    =   'False'
    AND A2EExcluded     =   'False'
    AND NOT EXISTS (
       SELECT * FROM A2E_UserProperties AS RGT
       WHERE LFT.A2EKey        =   RGT.A2EKey
       AND LFT.A2EDstMessageID =   RGT.A2ESrcMessageID
       AND LFT.A2ESrcMessageID =   RGT.A2EDstMessageID
       )
    " -ServerInstance "A2ESQLSERVER" | Out-File -FilePath "C:\sqltest.txt"

} 

# Option 1: Remove Bad Rows

if ($decision -eq 1) {


}

if ($decision -eq 2) {

Exit
}



# End Scripting