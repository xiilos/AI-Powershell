<#
.NAME
    Outlook Tools
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$OutlookTools_Menu               = New-Object system.Windows.Forms.Form
$OutlookTools_Menu.ClientSize    = New-Object System.Drawing.Point(247,344)
$OutlookTools_Menu.text          = "Outlook Tools"
$OutlookTools_Menu.TopMost       = $false
$OutlookTools_Menu.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$Rearm_Office                    = New-Object system.Windows.Forms.Label
$Rearm_Office.text               = "rearm Office"
$Rearm_Office.AutoSize           = $true
$Rearm_Office.width              = 150
$Rearm_Office.height             = 10
$Rearm_Office.location           = New-Object System.Drawing.Point(14,165)
$Rearm_Office.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$TT                              = New-Object system.Windows.Forms.ToolTip
$TT.ToolTipTitle                 = "Tool_Tip"

$Outlook_Install32               = New-Object system.Windows.Forms.Label
$Outlook_Install32.text          = "Install Outlook 365 x86"
$Outlook_Install32.AutoSize      = $true
$Outlook_Install32.width         = 150
$Outlook_Install32.height        = 10
$Outlook_Install32.location      = New-Object System.Drawing.Point(15,15)
$Outlook_Install32.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$Install_Outlook64               = New-Object system.Windows.Forms.Label
$Install_Outlook64.text          = "Install Outlook 365 x64"
$Install_Outlook64.AutoSize      = $true
$Install_Outlook64.width         = 150
$Install_Outlook64.height        = 10
$Install_Outlook64.location      = New-Object System.Drawing.Point(15,45)
$Install_Outlook64.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$Outlook_Updates                 = New-Object system.Windows.Forms.Label
$Outlook_Updates.text            = "Check for Outlook Updates"
$Outlook_Updates.AutoSize        = $true
$Outlook_Updates.width           = 150
$Outlook_Updates.height          = 10
$Outlook_Updates.location        = New-Object System.Drawing.Point(15,75)
$Outlook_Updates.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$Disable_OSC                     = New-Object system.Windows.Forms.Label
$Disable_OSC.text                = "Disable Outlook Social Connector"
$Disable_OSC.AutoSize            = $true
$Disable_OSC.width               = 150
$Disable_OSC.height              = 10
$Disable_OSC.location            = New-Object System.Drawing.Point(15,135)
$Disable_OSC.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$Outlook_Profile_Set             = New-Object system.Windows.Forms.Label
$Outlook_Profile_Set.text        = "Setup Outlook Profile"
$Outlook_Profile_Set.AutoSize    = $true
$Outlook_Profile_Set.width       = 150
$Outlook_Profile_Set.height      = 10
$Outlook_Profile_Set.location    = New-Object System.Drawing.Point(15,105)
$Outlook_Profile_Set.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$Bypass_O365                     = New-Object system.Windows.Forms.Label
$Bypass_O365.text                = "Bypass Office 365 (Hybrid)"
$Bypass_O365.AutoSize            = $true
$Bypass_O365.width               = 150
$Bypass_O365.height              = 10
$Bypass_O365.location            = New-Object System.Drawing.Point(15,195)
$Bypass_O365.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$DisableOutlookUpdates           = New-Object system.Windows.Forms.Label
$DisableOutlookUpdates.text      = "Disable Outlook Updates"
$DisableOutlookUpdates.AutoSize  = $true
$DisableOutlookUpdates.width     = 25
$DisableOutlookUpdates.height    = 10
$DisableOutlookUpdates.location  = New-Object System.Drawing.Point(15,225)
$DisableOutlookUpdates.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$TT.SetToolTip($Rearm_Office,'Re-arms Office Suit Trial Extension')
$TT.SetToolTip($Outlook_Install32,'Installs Outlook 365 32bit')
$TT.SetToolTip($Install_Outlook64,'Installs Outlook 365 64bit')
$TT.SetToolTip($Outlook_Updates,'Checks and runs Outlook 365 Updates')
$TT.SetToolTip($Disable_OSC,'Disable Outlook Social Connector Add-In')
$TT.SetToolTip($Outlook_Profile_Set,'Sets up the Outlook Profile for Add2Exchange')
$TT.SetToolTip($Bypass_O365,'Bypasses O365 Login for Hybrid Environments')
$TT.SetToolTip($DisableOutlookUpdates,'Disables Automatic Outlook Updates')
$OutlookTools_Menu.controls.AddRange(@($Rearm_Office,$Outlook_Install32,$Install_Outlook64,$Outlook_Updates,$Disable_OSC,$Outlook_Profile_Set,$Bypass_O365,$DidItBetter_logo,$DisableOutlookUpdates))


$Rearm_Office.Add_Click({Start-Process Powershell .\REARM_Office.ps1})
$Outlook_Install32.Add_Click({Start-Process Powershell .\Outlook_Installer.ps1})
$Install_Outlook64.Add_Click({Start-Process Powershell .\Outlook_Installer.ps1})
$Outlook_Updates.Add_Click({Start-Process Powershell .\Office_Updater.ps1})
$Disable_OSC.Add_Click({Start-Process Powershell .\OSC_Disable.bat})
$Outlook_Profile_Set.Add_Click({Start-Process Powershell .\Outlook_Profile_Set.ps1})
$Bypass_O365.Add_Click({Start-Process Powershell .\Bypass_AutoDiscover.ps1})
$DisableOutlookUpdates.Add_Click({Start-Process Powershell .\Disable_Outlook_Updates.ps1})

[void]$OutlookTools_Menu.ShowDialog()

# End Scripting