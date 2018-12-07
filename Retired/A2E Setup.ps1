<# 
.NAME
    A2E Menu
.SYNOPSIS
    Multiple scripts for Add2Exchange permissions and setup
.DESCRIPTION
    Menu for Powershell Scripts
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$A2EMenu                         = New-Object system.Windows.Forms.Form
$A2EMenu.ClientSize              = '285,546'
$A2EMenu.text                    = "Add2Exchange Setup"
$A2EMenu.TopMost                 = $false

$O365_onprem_permissions         = New-Object system.Windows.Forms.Button
$O365_onprem_permissions.text    = "O365 / Exchange On-Prem Permissions"
$O365_onprem_permissions.width   = 270
$O365_onprem_permissions.height  = 30
$O365_onprem_permissions.location  = New-Object System.Drawing.Point(5,20)
$O365_onprem_permissions.Font    = 'Microsoft Sans Serif,10,style=Bold'

$autologon                       = New-Object system.Windows.Forms.Button
$autologon.text                  = "Auto Logon"
$autologon.width                 = 270
$autologon.height                = 30
$autologon.location              = New-Object System.Drawing.Point(5,60)
$autologon.Font                  = 'Microsoft Sans Serif,10,style=Bold'

$Windows10Sanatize               = New-Object system.Windows.Forms.Button
$Windows10Sanatize.text          = "Sanitize Windows 10"
$Windows10Sanatize.width         = 270
$Windows10Sanatize.height        = 30
$Windows10Sanatize.location      = New-Object System.Drawing.Point(5,420)
$Windows10Sanatize.Font          = 'Microsoft Sans Serif,10,style=Bold'

$Add2ExchangeUpdate              = New-Object system.Windows.Forms.Button
$Add2ExchangeUpdate.text         = "Upgrade Add2Exchange"
$Add2ExchangeUpdate.width        = 270
$Add2ExchangeUpdate.height       = 30
$Add2ExchangeUpdate.location     = New-Object System.Drawing.Point(5,100)
$Add2ExchangeUpdate.Font         = 'Microsoft Sans Serif,10,style=Bold'

$PowershellUpgrade               = New-Object system.Windows.Forms.Button
$PowershellUpgrade.text          = "Upgrade Powershell"
$PowershellUpgrade.width         = 270
$PowershellUpgrade.height        = 30
$PowershellUpgrade.location      = New-Object System.Drawing.Point(5,138)
$PowershellUpgrade.Font          = 'Microsoft Sans Serif,10,style=Bold'

$GPResults                       = New-Object system.Windows.Forms.Button
$GPResults.text                  = "Check Group Policy"
$GPResults.width                 = 270
$GPResults.height                = 30
$GPResults.location              = New-Object System.Drawing.Point(5,180)
$GPResults.Font                  = 'Microsoft Sans Serif,10,style=Bold'

$RegistryFavorites               = New-Object system.Windows.Forms.Button
$RegistryFavorites.text          = "Add Registry Favorites"
$RegistryFavorites.width         = 270
$RegistryFavorites.height        = 30
$RegistryFavorites.location      = New-Object System.Drawing.Point(5,220)
$RegistryFavorites.Font          = 'Microsoft Sans Serif,10,style=Bold'

$DisableUAC                      = New-Object system.Windows.Forms.Button
$DisableUAC.text                 = "Disable UAC"
$DisableUAC.width                = 270
$DisableUAC.height               = 30
$DisableUAC.location             = New-Object System.Drawing.Point(5,260)
$DisableUAC.Font                 = 'Microsoft Sans Serif,10,style=Bold'

$OutlookAddinsdisable            = New-Object system.Windows.Forms.Button
$OutlookAddinsdisable.text       = "Disable Outlook Add-ins"
$OutlookAddinsdisable.width      = 270
$OutlookAddinsdisable.height     = 30
$OutlookAddinsdisable.location   = New-Object System.Drawing.Point(5,300)
$OutlookAddinsdisable.Font       = 'Microsoft Sans Serif,10,style=Bold'

$Outlook365Install               = New-Object system.Windows.Forms.Button
$Outlook365Install.text          = "Install Outlook 365"
$Outlook365Install.width         = 270
$Outlook365Install.height        = 30
$Outlook365Install.location      = New-Object System.Drawing.Point(5,340)
$Outlook365Install.Font          = 'Microsoft Sans Serif,10,style=Bold'

$ExportLicense_Profile1          = New-Object system.Windows.Forms.Button
$ExportLicense_Profile1.text     = "Export License Info and Profile 1"
$ExportLicense_Profile1.width    = 270
$ExportLicense_Profile1.height   = 30
$ExportLicense_Profile1.location  = New-Object System.Drawing.Point(5,380)
$ExportLicense_Profile1.Font     = 'Microsoft Sans Serif,10,style=Bold'

$ThatWasEasy                     = New-Object system.Windows.Forms.Button
$ThatWasEasy.text                = "Make it Easy"
$ThatWasEasy.width               = 270
$ThatWasEasy.height              = 30
$ThatWasEasy.location            = New-Object System.Drawing.Point(5,500)
$ThatWasEasy.Font                = 'Microsoft Sans Serif,16,style=Bold'

$redemption                      = New-Object system.Windows.Forms.Button
$redemption.text                 = "Legacy Redemption Un-Register"
$redemption.width                = 270
$redemption.height               = 30
$redemption.location             = New-Object System.Drawing.Point(5,460)
$redemption.Font                 = 'Microsoft Sans Serif,10,style=Bold'

$A2EMenu.controls.AddRange(@($O365_onprem_permissions,$autologon,$Windows10Sanatize,$Add2ExchangeUpdate,$PowershellUpgrade,$GPResults,$RegistryFavorites,$DisableUAC,$OutlookAddinsdisable,$Outlook365Install,$ExportLicense_Profile1,$ThatWasEasy,$redemption))

#region gui events {
$O365_onprem_permissions.Add_Click({PermissionsOnPremOrO365Combined})
$autologon.Add_Click({AutoLogin})
$Add2ExchangeUpdate.Add_Click({A2E_Updates})
$PowershellUpgrade.Add_Click({Powershell_MSOnline_Update})
$GPResults.Add_Click({GP_Results})
$RegistryFavorites.Add_Click({Registry_Favorites})
$DisableUAC.Add_Click({Disable_UAC})
$OutlookAddinsdisable.Add_Click({Remove_Outlook_Add_ins})
$Outlook365Install.Add_Click({     })
$ExportLicense_Profile1.Add_Click({Export_License_and_Profile1})
$Windows10Sanatize.Add_Click({SanitizeWindows10})
$redemption.Add_Click({Redemption})
$ThatWasEasy.Add_Click({1ClickSetup})
#endregion events }

#endregion GUI }


#Functions start below here

Function PermissionsOnPremOrO365Combined {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
& $setup\PermissionsOnPremOrO365Combined.ps1
}


Function AutoLogin {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\AutoLogin.ps1
}


Function A2E_Updates {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\A2E_Updates.ps1

}


Function Powershell_MSOnline_Update {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Powershell_MSOnline_Update.ps1

}

Function GP_Results {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\GP_Results.ps1

}


Function Registry_Favorites {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Registry_Favorites.ps1
}

Function Disable_UAC {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Disable_UAC.ps1

}


Function Remove_Outlook_Add_ins {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\ReEnable_Outlook_Add_ins.ps1

}

Function ReEnable_Outlook_Add_ins {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Remove_Outlook_Add_ins.ps1
}

Function Export_License_and_Profile1 {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Export_License_and_Profile1.ps1

}

Function VirginizeWindows10 {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\VirginizeWindows10.ps1

}


Function Redemption {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\Redemption.ps1
}

Function 1ClickSetup {
    $scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    & $setup\1ClickSetup.ps1
}


[void]$A2EMenu.ShowDialog()