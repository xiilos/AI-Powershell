if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
  }
  
  
  #Execution Policy
  
  Set-ExecutionPolicy -ExecutionPolicy Bypass
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  #Variables
  
  $Exchangename = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Name.txt"
  $ServiceAccount = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
  $Username = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
  $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" | convertto-securestring
  
  # Script #
  
  Try {
  
    $Cred = New-Object -typename System.Management.Automation.PSCredential `
      -Argumentlist $Username, $Password
  
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangename/PowerShell/ -Authentication Kerberos -Credential $Cred
    Import-PSSession $Session -DisableNameChecking
    Set-ADServerSettings -ViewEntireForest $true
  
  
    Get-Mailbox -Resultsize Unlimited | Add-MailboxPermission -User $ServiceAccount -AccessRights FullAccess -InheritanceType all -AutoMapping:$false -confirm:$false
    Get-Mailbox -Resultsize Unlimited | Where-Object { $_.WhenCreated -ge ((Get-Date).Adddays(-90)) } | Add-MailboxPermission -User $ServiceAccount -AccessRights 'FullAccess' -InheritanceType all -AutoMapping:$false -Confirm:$false
  }
  
  Catch {
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Get-PSSession | Remove-PSSession
    Exit
  }
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to All Users On-Premise Succesfully."
  
  Get-PSSession | Remove-PSSession
  Exit
  
  # End Scripting