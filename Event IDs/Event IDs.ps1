Logging Errors:

-EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to All Users On-Premise Succesfully."


-EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"

-EventID 10002 -EntryType SuccessAudit -Message "Add2Exchange Succesfully Backed up A2E Database."

-EventID 10003 -EntryType FailureAudit -Message "Add2Exchange Console is Open and cannot backup the A2E SQL Databse $_.Exception.Message"

-EventID 10004 -EntryType Information -Message "Add2Exchange did not succesfully backup the DB. Starting up services."

-EventID 10020 -EntryType FailureAudit -Message "SQL Transaction Log Trim failure $_.Exception.Message"

-EventID 10021 -EntryType FailureAudit -Message "Add2Exchange SQL Transaction Log Trimmed Succesfully"


Try {
}



  Catch {
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10001 -EntryType FailureAudit -Message "$_.Exception.Message"
    Get-PSSession | Remove-PSSession
    Exit
  }
  
    Write-EventLog -LogName "Add2Exchange" -Source "Add2Exchange" -EventID 10000 -EntryType SuccessAudit -Message "Add2Exchange PowerShell Added Permissions to All Users On-Premise Succesfully."