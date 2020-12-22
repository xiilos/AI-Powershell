if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

# Script #

#Send Mail Settings
$SendEmail = $false                    # = $true if you want to enable send report to e-mail (SMTP send)
$EmailTo   = 'kouretas.k@gmail.com'              #user@domain.something (for multiple users use "User01 &lt;user01@example.com&gt;" ,"User02 &lt;user02@example.com&gt;" )
$EmailFrom = 'dkouretas@diditbetter.com'   #matthew@domain 
$EmailSMTP = 'smtp.ex16.diditbetter.com' #smtp server adress, DNS hostname.




# Send e-mail with reports as attachments
if ($SendEmail -eq $true) {
  $EmailSubject = "Backup Email $(get-date -format MM.yyyy)"
  $EmailBody = "Backup Script $(get-date -format MM.yyyy) (last Month).`nYours sincerely `Matthew - SYSTEM ADMINISTRATOR"
  Logging "INFO" "Sending e-mail to $EmailTo from $EmailFrom (SMTPServer = $EmailSMTP) "
  ### the attachment is $log 
  Send-MailMessage -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody -SmtpServer $EmailSMTP -attachment $Log 
}

$Log = "C:\zlibrary\TEMP\Logs\Event_Log.evtx"

  Send-MailMessage -To "kouretas.k@gmail.com" -From "dkouretas@diditbetter.com" -Subject "Add2Exchange Logs" -Body "This is just a test for email" -SmtpServer "192.168.0.11" -port 25 -attachment $Log



  $yesterday = (Get-Date).AddHours(-24)
  $ErrWarn4App = Get-WinEvent -FilterHashTable @{LogName='Add2Exchange'; Level=2,3; StartTime=$yesterday} -ErrorAction SilentlyContinue | Select-Object TimeCreated,LogName,ProviderName,Id,LevelDisplayName,Message
  $ErrWarn4App | Sort-Object TimeCreated | Format-Table -AutoSize 



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting