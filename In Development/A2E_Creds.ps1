if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}


<# Notes*
1. Reg key Service Account (Put it there and Get it there)
2. Outlook Profile Password Create (Put it there and Get it there)
3. Current/Local user logged in as. domain\ (Put it there and Get it there)
4. Current/Local Account Password (Put it there and Get it there)

password validation

#>

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#1. Getting the Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount" -ErrorAction SilentlyContinue
Write-Host "The Service Account Name is $ServiceAccount"


#Setting the Service Account Name
$ServiceAccount = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount" -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount" -Value "$ServiceAccount" -ErrorAction SilentlyContinue

<#If we create a seperate file with the service account name then we have to 
- Create the file
- Read the file
- Import the contents of file into werever we are trying to go #>

#Create
Read-Host "What is the Service Account Name?" | out-file ".\ServiceAccount.txt"

#Read
$ServiceAccount = Get-Content ".\ServiceAccount.txt"

#Write/Import
$ServiceAccount = Get-Content ".\ServiceAccount.txt"
Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "ServiceAccount" -Value "$ServiceAccount" -ErrorAction SilentlyContinue



#2. Getting Outlook Profile Password

#Since the Password is Encrypted in the registry, we will have to store the password somewere using windows bit lock

#Create
Read-Host "Global Admin Password or Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\ServerPass.txt"

#Read
$Password = Get-Content ".\ServerPass.txt" | convertto-securestring
Write-Host "The password is $Password"

#Write/Import
$Password = Get-Content ".\ServerPass.txt" | convertto-securestring
Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange" -Name "serviceAccountPwd" -Value "$Password" -ErrorAction SilentlyContinue


#3. Getting the current User and Domain name
#This will get both current username and domain in this format:  Domain\Username

$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
Write-Host "$Username"

#Sending this information to a file
#Create
$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name | out-file ".\Username.txt"

#Read
$Username = Get-Content ".\Username.txt"
Write-Host "The domain and username is $Username"

#Write/Import
$Username = Get-Content ".\Username.txt"
$Password = Get-Content ".\ServerPass.txt" | convertto-securestring

#Write it to the Service
$SVC = Get-WmiObject win32_service -Filter "Name='Add2Exchange Service'"
$SVC.StopService();
$Result = $SVC.Change($Null, $Null, $Null, $Null, $Null, $Null, $Username, "$Password")
If ($Result.ReturnValue -eq '0') { Write-Host "Add2Exchange Service Username and Password Has Been Succsefully Updated" -ForegroundColor Green } Else { Write-Host "Error: $Result" }


#4. Getting the current local account password and using it
#We can do this in 2 ways. 1. We can store the password off of windows credential manager and use th same method as Step 2. Or we can take it
#a step further and use windows credential to add/edit/modify through powershell. 

#Install the credential manager module in powershell
Install-Module -Name 'CredentialManager'

#Creating a credential and adding it into Windows Credential Manager
#Create it
$Test = @{
    Target   = 'A2E_test'
    UserName = 'zAdd2Exchange'
    Password = '1234'
    Comment  = 'zAdd2Exchange username and password'
    Persist  = 'LocalMachine'
}
#Store it
New-StoredCredential @Test

#Call It
Get-StoredCredential -target 'A2E_test'

#or

Get-StoredCredential -target 'a2e_test' -AsCredentialObject

#This will show the password as well

#Finally to get a list of all username and password on the box

Get-StoredCredential -AsCredentialObject | out-file c:\passwords.txt


# End Scripting