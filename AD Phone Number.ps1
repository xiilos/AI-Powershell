if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #


Get-ADUser -SearchBase "CN=Users,DC=MBX,DC=local" -Filter "telephoneNumber -like '+44*'" -Properties telephoneNumber |
    ForEach-Object {
        $PhoneNumberRaw = $_.telephoneNumber -replace '^0' -replace '^\44' -replace '\s' -as [LONG]
        $newPhoneNumber = if ($PhoneNumberRaw -match '^+44') {
                            "+011 {0:# #### ####}" -f $PhoneNumberRaw
                        }
                        else {
                            "+011 {0:### ### ###}" -f $PhoneNumberRaw
                        }
        [PSCustomObject]@{
            Name = $_.Name
            TelephoneNumber = $newPhoneNumber
        }
    $Users = Get-ADUser -SearchBase "CN=Users,DC=MBX,DC=local" -Filter "telephoneNumber -like '+44*'" -Properties telephoneNumber, distinguishedName
        ForEach ($User In $Users)
    {
        $DN = $User.distinguishedName
        Set-ADUser -Identity $DN -OfficePhone $newPhoneNumber
    }
}
Pause
Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting