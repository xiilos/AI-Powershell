if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Test for Upgrade Eligibility
$LicenseKeyAExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyAExpiry" -ErrorAction SilentlyContinue
$LicenseKeyCExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyCExpiry" -ErrorAction SilentlyContinue
$LicenseKeyEExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyEExpiry" -ErrorAction SilentlyContinue
$LicenseKeyGExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyGExpiry" -ErrorAction SilentlyContinue
$LicenseKeyMExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyMExpiry" -ErrorAction SilentlyContinue
$LicenseKeyNExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyNExpiry" -ErrorAction SilentlyContinue
$LicenseKeyOExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyOExpiry" -ErrorAction SilentlyContinue
$LicenseKeyPExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyPExpiry" -ErrorAction SilentlyContinue
$LicenseKeyTExpiry = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange\Profile 1" -Name "LicenseKeyTExpiry" -ErrorAction SilentlyContinue

$Today = Get-Date

#License Varify

#Calendars--------------

$Calendars = if ($LicenseKeyAExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyAExpiry -and $LicenseKeyAExpiry -notlike "") {
    "$LicenseKeyAExpiry !!EXPIRED!!" 
}

Else {
    "$LicenseKeyAExpiry"
}

#Contacts--------------

$Contacts = if ($LicenseKeyCExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyCExpiry -and $LicenseKeyCExpiry -notlike "") {
    "$LicenseKeyCExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyCExpiry"
}

#RGM--------------

$RGM = if ($LicenseKeyEExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyEExpiry -and $LicenseKeyEExpiry -notlike "") {
    "$LicenseKeyEExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyEExpiry"
}

#GAL--------------

$GAL = if ($LicenseKeyGExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyGExpiry -and $LicenseKeyGExpiry -notlike "") {
    "$LicenseKeyGExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyGExpiry"
}


#MAIL--------------

$Mail = if ($LicenseKeyMExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyMExpiry -and $LicenseKeyMExpiry -notlike "") {
    "$LicenseKeyMExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyMExpiry"
}


#EMAIL--------------

$Email = if ($LicenseKeyNExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyNExpiry -and $LicenseKeyNExpiry -notlike "") {
    "$LicenseKeyNExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyNExpiry"
}


#Notes--------------

$Notes = if ($LicenseKeyOExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyOExpiry -and $LicenseKeyOExpiry -notlike "") {
    "$LicenseKeyOExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyOExpiry"
}


#Posts--------------

$Posts = if ($LicenseKeyPExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyPExpiry -and $LicenseKeyPExpiry -notlike "") {
    "$LicenseKeyPExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyPExpiry"
}


#Tasks--------------

$Tasks = if ($LicenseKeyTExpiry -eq "") {
    "Not Licensed or in Trial"
}

Elseif ($Today -ge $LicenseKeyTExpiry -and $LicenseKeyTExpiry -notlike "") {
    "$LicenseKeyTExpiry !!EXPIRED!!"
}

Else {
    "$LicenseKeyTExpiry"
}


$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
$answer = $wshell.Popup("Please Review your Add2Exchange License Expiration Dates below.

If your Free Upgrade Period (Software Assurance) has passed and they are expired, you will need to renew the software to continue using it.
If you have already renewed or in that process, please Press Continue.

Your Software Mainenance Free Upgrade Periods

$Calendars for Calendars Synchronization
$Contacts for Contacts Synchronization
$RGM  for the Relationship Group Manager 
$GAL for GAL Synchronization 
$Notes for Notes Synchronization 
$Posts  for Posts Synchronization 
$Tasks for Tasks Synchronization 
$Mail for Email Synchronization
$Email  for Confidential Email Notifier

$Today is today’s date, and any keys for your active functioning modules should expire AFTER this date.  

Click OK to continue with your upgrade, or Cancel to Quit.

NOTE* Upgrading Add2Exchange with expired keys and outside Software Assurance will stop synchronization until a renewal is purchased and new keys are issued.
Also, if the service account can receive email, after you purchase, the keys will automatically be applied, usually without intervention.

", 0, "ATTENTION!! Add2Exchange Licensing", 0 + 1)
if ($answer -eq 2) { Break }




#Stop Menu Process
Stop-Process -Name "DidItBetter Support Menu" -Force -ErrorAction SilentlyContinue

#Test for FTP

try {
    $FTP = New-Object System.Net.Sockets.TcpClient("ftp.diditbetter.com", 21)
    $FTP.Close()
    Write-Host "Connectivity OK."
}
catch {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    $wshell.Popup("No FTP Access... Taking you to Downloads.... Click OK or Cancel to Quit.", 0, "ATTENTION!!", 0 + 1)
    Start-Process "http://support.diditbetter.com/Secure/Login.aspx?returnurl=/downloads.aspx"
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#Stop Add2Exchange Service

Write-Host "Stopping Add2Exchange Service"
Stop-Service -Name "Add2Exchange Service"
Start-Sleep -s 5
Write-Host "Done"

#Stop The Add2Exchange Agent

Write-Host "Stopping the Agent. Please Wait."
Start-Sleep -s 10
$Agent = Get-Process "Add2Exchange Agent" -ErrorAction SilentlyContinue
if ($Agent) {
    Write-Host "Waiting for Agent to Exit"
    Start-Sleep -s 10
    if (!$Agent.HasExited) {
        $Agent | Stop-Process -Force
    }
}


#Remove Add2Exchange

Write-Host "Removing Add2Exchange"
Write-Host "Please Wait...."
$Program = Get-WmiObject -Class Win32_Product -Filter "Name = 'Add2Exchange'"
$Program.Uninstall()
Write-Host "Done"

#Create zLibrary

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\Add2Exchange Upgrades"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "Add2Exchange Upgrades Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\Add2Exchange Upgrades"
}

#Downloading Add2Exchange

Write-Host "Downloading Add2Exchange"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/A2E-Enterprise/Upgrades/a2e-enterprise_upgrade.exe"
$Output = "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Exchange

Write-Host "Unpacking Add2exchange"
Write-Host "please Wait....."
Push-Location "c:\zlibrary\Add2Exchange Upgrades"
Start-Process "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe" -wait
Start-Sleep -Seconds 2
Write-Host "Done"

#Installing Add2Exchange
Do {
    Write-Host "Installing Add2Exchange"
    $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Push-Location $Location
    Start-Process msiexec.exe -Wait -ArgumentList '/I "Add2Exchange_Upgrade.msi" /quiet' -ErrorAction Inquire -ErrorVariable InstallError;
    Write-Host "Finished...Upgrade Complete"

    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))



#Setting the Service Account Password
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password)
$SAP = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)

$SVC = Get-WmiObject win32_service -Filter "Name='Add2Exchange Service'"
$SVC.StopService();
$Result = $SVC.Change($Null, $Null, $Null, $Null, $Null, $Null, $Null, "$SAP")
If ($Result.ReturnValue -eq '0') { Write-Host "Add2Exchange Service Password Has Been Succsefully Updated" -ForegroundColor Green } Else { Write-Host "Error: $Result" }



#Setting the Add2Exchange Service to Delayed Start
Write-Host "Setting up Add2Exchange Service to Delayed Start"
sc.exe config "Add2Exchange Service" start= delayed-auto
Write-Host "Done"

#Starting Add2Exchange Console
Write-Host "Starting the Add2Exchange Console"
$Install = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "InstallLocation" -ErrorAction SilentlyContinue
Push-Location $Install
Start-Process "./Console/Add2Exchange Console.exe"

#Start-Service -Name "Add2Exchange Service"
Write-Host "Done"



Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting
# SIG # Begin signature block
# MIIWHAYJKoZIhvcNAQcCoIIWDTCCFgkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4Fs+37PGFCjehAddGB4542S2
# r2qgghCOMIIDDjCCAfagAwIBAgIQKrltcrcv9YBOoOhyXdcOLjANBgkqhkiG9w0B
# AQsFADAfMR0wGwYDVQQDDBREaWRJVEJldHRlciBTb2Z0d2FyZTAeFw0yMjA0MDUy
# MDIyMjdaFw0yMzA0MDUyMDQyMjdaMB8xHTAbBgNVBAMMFERpZElUQmV0dGVyIFNv
# ZnR3YXJlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqhtZjlXCb+yt
# IcrwXJFlLYI9gRbc2eLfKSKs66/ES8Lp6opJaoN0qQ1frMWGpkvVsSVN9J7mL0cE
# OZhyfSOQWnBMY+FWqg0/gdh7kF+4ZmTTQ/CL43G3J1Hm/E8545gYclvmafo4R7JQ
# ttRQprGZkbp/GvjUY36TUk/SD0ZzQeLU0wqcXx5uXLQF3O4arshaSNOAWUsaPDXi
# FhBToTUuHA1ZwHxgfQ+uLM8nsMMok35S4Y6ssShf851To6ssFAxf/pypxmlo1xBC
# aIf6RtgRnkOt3rFyLKyHoQiYP0qG//HMx9Sivp/3f2u72fELHuYf2V0U6v9DG28h
# mFNGaVgebQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwHQYDVR0OBBYEFMzlZDrV0mbfaWpLto1xsg4x8KcRMA0GCSqGSIb3DQEB
# CwUAA4IBAQBjs0HeAWYEjfZozljk5HcbJZ9h1gCMDe8Iq3r5L0l+1tFBou/L011z
# PnuFOjigVMUnVHhtoktZkrsdnOiPGl6txlcSyG/FCD3CM4i8Pw7tv2N6uAd79W+1
# +EvnsQ+puh7UYeMyfXOga0gFE5yERjrYnvQn3xh9ELG1m1oT5aM3gcEcR8pglgc1
# r72a9rl+nFVSEVDNKsyUT01l3kKHHg1rXOcSOa+t6V5Ma4W9o+JJYD3fBNP0rrUY
# U6WRebv7G/6o6Buctl524icTrSZg8RLeA2QdDo03NKRgksFeegrg6O0bz1hUTgJb
# rJffWfU/d34gSU14MgnJ2Q6ZQhoyjbZ8MIIGrjCCBJagAwIBAgIQBzY3tyRUfNhH
# rP0oZipeWzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcw
# MzIyMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIElu
# Yy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYg
# VGltZVN0YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# xoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo
# +n3znIkLf50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY3+vaPcQX
# f6sZKz5C3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8
# ald68Dd5n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71ORJn7w6l
# Y2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW8LAxE6lX
# KZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU
# 2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2
# RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR
# 8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI/rOP015L
# dhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNY
# CQEoAA6EVO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0wggFZMBIG
# A1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WMaiCPnshv
# MB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIB
# hjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQw
# QwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZI
# AYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBgu
# EE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFn
# zbYSlm/EUExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/t
# YLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/KFSJ8NWK
# cXZl2szwcqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7
# pp1yr8THwcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkP
# lM05et3/JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRxktcma+Q4
# c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kad
# dSweJywm228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrdVcA6KYaw
# mKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL
# 5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QVSucTDh3b
# NzgaoSv27dZ8/DCCBsYwggSuoAMCAQICEAp6SoieyZlCkAZjOE2Gl50wDQYJKoZI
# hvcNAQELBQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRp
# bWVTdGFtcGluZyBDQTAeFw0yMjAzMjkwMDAwMDBaFw0zMzAzMTQyMzU5NTlaMEwx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjEkMCIGA1UEAxMb
# RGlnaUNlcnQgVGltZXN0YW1wIDIwMjIgLSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAuSqWI6ZcvF/WSfAVghj0M+7MXGzj4CUu0jHkPECu+6vE43hd
# flw26vUljUOjges4Y/k8iGnePNIwUQ0xB7pGbumjS0joiUF/DbLW+YTxmD4LvwqE
# EnFsoWImAdPOw2z9rDt+3Cocqb0wxhbY2rzrsvGD0Z/NCcW5QWpFQiNBWvhg02Us
# Pn5evZan8Pyx9PQoz0J5HzvHkwdoaOVENFJfD1De1FksRHTAMkcZW+KYLo/Qyj//
# xmfPPJOVToTpdhiYmREUxSsMoDPbTSSF6IKU4S8D7n+FAsmG4dUYFLcERfPgOL2i
# vXpxmOwV5/0u7NKbAIqsHY07gGj+0FmYJs7g7a5/KC7CnuALS8gI0TK7g/ojPNn/
# 0oy790Mj3+fDWgVifnAs5SuyPWPqyK6BIGtDich+X7Aa3Rm9n3RBCq+5jgnTdKEv
# sFR2wZBPlOyGYf/bES+SAzDOMLeLD11Es0MdI1DNkdcvnfv8zbHBp8QOxO9APhk6
# AtQxqWmgSfl14ZvoaORqDI/r5LEhe4ZnWH5/H+gr5BSyFtaBocraMJBr7m91wLA2
# JrIIO/+9vn9sExjfxm2keUmti39hhwVo99Rw40KV6J67m0uy4rZBPeevpxooya1h
# sKBBGBlO7UebYZXtPgthWuo+epiSUc0/yUTngIspQnL3ebLdhOon7v59emsCAwEA
# AaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB
# /wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwH
# ATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUjWS3
# iSH+VlhEhGGn6m8cNo/drw0wWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRp
# bWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEADS0jdKbR9fjqS5k/
# AeT2DOSvFp3Zs4yXgimcQ28BLas4tXARv4QZiz9d5YZPvpM63io5WjlO2IRZpbwb
# mKrobO/RSGkZOFvPiTkdcHDZTt8jImzV3/ZZy6HC6kx2yqHcoSuWuJtVqRprfdH1
# AglPgtalc4jEmIDf7kmVt7PMxafuDuHvHjiKn+8RyTFKWLbfOHzL+lz35FO/bgp8
# ftfemNUpZYkPopzAZfQBImXH6l50pls1klB89Bemh2RPPkaJFmMga8vye9A140pw
# SKm25x1gvQQiFSVwBnKpRDtpRxHT7unHoD5PELkwNuTzqmkJqIt+ZKJllBH7bjLx
# 9bs4rc3AkxHVMnhKSzcqTPNc3LaFwLtwMFV41pj+VG1/calIGnjdRncuG3rAM4r4
# SiiMEqhzzy350yPynhngDZQooOvbGlGglYKOKGukzp123qlzqkhqWUOuX+r4DwZC
# nd8GaJb+KqB0W2Nm3mssuHiqTXBt8CzxBxV+NbTmtQyimaXXFWs1DoXW4CzM4Awk
# uHxSCx6ZfO/IyMWMWGmvqz3hz8x9Fa4Uv4px38qXsdhH6hyF4EVOEhwUKVjMb9N/
# y77BDkpvIJyu2XMyWQjnLZKhGhH+MpimXSuX4IvTnMxttQ2uR2M4RxdbbxPaahBu
# H0m3RFu0CAqHWlkEdhGhp3cCExwxggT4MIIE9AIBATAzMB8xHTAbBgNVBAMMFERp
# ZElUQmV0dGVyIFNvZnR3YXJlAhAquW1yty/1gE6g6HJd1w4uMAkGBSsOAwIaBQCg
# eDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJ
# BDEWBBSeCYNaNPa2tBRRmwV/+WNiXbMhEjANBgkqhkiG9w0BAQEFAASCAQB0JSkl
# Y7OJcOSYjXv2ltwasIZzDzrzALXoOY4S2MUn0+wVGpM5RDUyrzy9vm4NuSPto2Ro
# 7tmxVf+5Ku2eENjoV3TWvcjtQcvAzndTF3xTFofLdKclUwxjeRAPZyLFaza/QOak
# gx0C0gBtsfz0C/gewh7nOOkJPMkYS7dZvYLj3VIYQ5BA4y7grRtqWZ6JJwcss4/t
# 6MI6jbU6gE+38YQn2x4Coyq6iwyRoUnhGdUt8Q/TvI5xY2pRgXbMMnS+MX2u6Gck
# o3YokDZwloQXzl4B92j9VoHvzGzbg7HKUSyZxc8+DeuaVcKXjUWUXM7mDTRhywE+
# Pm+9OJdGAwdd+X+UoYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJE
# aWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBD
# QQIQCnpKiJ7JmUKQBmM4TYaXnTANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIyMDQwNTIwMzIzMlowLwYJ
# KoZIhvcNAQkEMSIEID+SRON+YeKSEBMG6HxGadNtRoHK8FNwVdZ064FVDuMuMA0G
# CSqGSIb3DQEBAQUABIICALRXSRn4Uz1W8j0e5tXNZmts3Y0XYUygh5iNW//srmif
# AC51eS83BM1VLEtHMFjuoqcecyuypNP+YtpKh+VFaD26hFM0hmsr1Oy8AXZ1wIZj
# tp1h4TG4Cm4Z3qDQBUVSkszxU3y3mmv1rc6LfkiyXhjmIaPoE32oNlc+OpUjDxWR
# WzqCzNWSTtkoBYXFurw6Erru+2Uu8BllrZdw0uCdfEwDvaZZwZ2J6rdlasLTfly8
# 3fA1tz4pH8hPOnG3iii0DPStI4m4pgjGv5H8XPf4UK2+HOlKbNT9TwiPX5abD8CG
# 1O+qwnkgrWIV6VehB4T0WMiMDauod0Fi7ARTpD1OEmM2zU6SvMngv2vEi3tFIqf7
# l+t3PN7CkoIINSDZ3eC9ZXTgJ7uA/eLRbokHIn7hnyg9gvuGyHaMcavQQ50ZY0sv
# 0NyxEOK3tuMFq1Cshu3Hh9V7Mry0viX1XhOf/rzEdlSdwtp6hmvoxvnTJMJrWdaE
# BtV9qyMh+K85w44VQC+ZNhx62fBQsoPKLmwV/6p9ghzDCCBA9n+Qz0Kta2RFdCep
# BevO7v/HFuHeXUXGkQV81j3AE9w0pzo9gJx81QSHp6N1DyGKLVqD279tLts1QaTy
# 7Ld6eOhxHBuuuUtpZMbWXoquSlMpA+CBSE0OZMBFo8mW5XoKxPzrmrL/rB+ocV/6
# SIG # End signature block
