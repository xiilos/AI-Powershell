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
$A2EMenu.ClientSize              = '285,525'
$A2EMenu.text                    = "Add2Exchange Setup"
$A2EMenu.TopMost                 = $false

$O365_onprem_permissions         = New-Object system.Windows.Forms.Button
$O365_onprem_permissions.text    = "Office 365 & Exchange Permissions"
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

$Windows10Virgin                 = New-Object system.Windows.Forms.Button
$Windows10Virgin.text            = "Virginize Windows 10"
$Windows10Virgin.width           = 270
$Windows10Virgin.height          = 30
$Windows10Virgin.location        = New-Object System.Drawing.Point(5,420)
$Windows10Virgin.Font            = 'Microsoft Sans Serif,10,style=Bold'

$Add2ExchangeUpdate              = New-Object system.Windows.Forms.Button
$Add2ExchangeUpdate.text         = "Upgrade Add2Exchange"
$Add2ExchangeUpdate.width        = 270
$Add2ExchangeUpdate.height       = 30
$Add2ExchangeUpdate.location     = New-Object System.Drawing.Point(5,100)
$Add2ExchangeUpdate.Font         = 'Microsoft Sans Serif,10,style=Bold'

$PowershellUpdate                = New-Object system.Windows.Forms.Button
$PowershellUpdate.text           = "MSOnline and PS Update"
$PowershellUpdate.width          = 270
$PowershellUpdate.height         = 30
$PowershellUpdate.location       = New-Object System.Drawing.Point(5,140)
$PowershellUpdate.Font           = 'Microsoft Sans Serif,10,style=Bold'

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

$outlookAddinsEnable             = New-Object system.Windows.Forms.Button
$outlookAddinsEnable.text        = "Enable Outlook Add-ins"
$outlookAddinsEnable.width       = 270
$outlookAddinsEnable.height      = 30
$outlookAddinsEnable.location    = New-Object System.Drawing.Point(5,340)
$outlookAddinsEnable.Font        = 'Microsoft Sans Serif,10,style=Bold'

$ExportLicense_Profile1          = New-Object system.Windows.Forms.Button
$ExportLicense_Profile1.text     = "Export License Info and Profile 1"
$ExportLicense_Profile1.width    = 270
$ExportLicense_Profile1.height   = 30
$ExportLicense_Profile1.location  = New-Object System.Drawing.Point(5,380)
$ExportLicense_Profile1.Font     = 'Microsoft Sans Serif,10,style=Bold'

$ThatWasEasy                     = New-Object system.Windows.Forms.Button
$ThatWasEasy.text                = "Easy 1 Click Setup"
$ThatWasEasy.width               = 270
$ThatWasEasy.height              = 30
$ThatWasEasy.location            = New-Object System.Drawing.Point(5,460)
$ThatWasEasy.Font                = 'Microsoft Sans Serif,16,style=Bold'

$A2EMenu.controls.AddRange(@($O365_onprem_permissions,$autologon,$Windows10Virgin,$Add2ExchangeUpdate,$PowershellUpdate,$GPResults,$RegistryFavorites,$DisableUAC,$OutlookAddinsdisable,$outlookAddinsEnable,$ExportLicense_Profile1,$ThatWasEasy))

#region gui events {
$O365_onprem_permissions.Add_Click({PermissionsOnPremOrO365Combined})
$autologon.Add_Click({AutoLogin})
$Add2ExchangeUpdate.Add_Click({A2E_Updates})
$PowershellUpdate.Add_Click({Powershell_MSOnline_Update})
$GPResults.Add_Click({GP_Results})
$RegistryFavorites.Add_Click({Registry_Favorites})
$DisableUAC.Add_Click({Disable_UAC})
$OutlookAddinsdisable.Add_Click({Remove_Outlook_Add_ins})
$outlookAddinsEnable.Add_Click({ReEnable_Outlook_Add_ins})
$ExportLicense_Profile1.Add_Click({Export_License_and_Profile1})
$Windows10Virgin.Add_Click({VirginizeWindows10})
$ThatWasEasy.Add_Click({1ClickSetup})
#endregion events }

#endregion GUI }


#Functions start below here
Function PermissionsOnPremOrO365Combined {

}


Function AutoLogin {


}


Function A2E_Updates {


}


Function Powershell_MSOnline_Update {


}

Function GP_Results {


}


Function Registry_Favorites {


}

Function Disable_UAC {


}


Function Remove_Outlook_Add_ins {


}

Function ReEnable_Outlook_Add_ins {

}

Function Export_License_and_Profile1 {


}

Function VirginizeWindows10 {


}

Function 1ClickSetup {


}



[void]$A2EMenu.ShowDialog()