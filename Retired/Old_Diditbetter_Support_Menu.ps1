<#
.NAME
    A2EMenu
.DESCRIPTION
    Add2Exchange Menu
#>

Set-ExecutionPolicy -ExecutionPolicy Bypass

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Menu               = New-Object system.Windows.Forms.Form
$Add2Exchange_Menu.ClientSize    = '1034,465'
$Add2Exchange_Menu.text          = "DiditBetter Software Add2Exchange Setup"
$Add2Exchange_Menu.BackColor     = "#ffffff"
$Add2Exchange_Menu.TopMost       = $false

$UpgradeAdd2Exchange             = New-Object system.Windows.Forms.Button
$UpgradeAdd2Exchange.text        = "Upgrade Add2Exchange"
$UpgradeAdd2Exchange.width       = 250
$UpgradeAdd2Exchange.height      = 30
$UpgradeAdd2Exchange.location    = New-Object System.Drawing.Point(15,58)
$UpgradeAdd2Exchange.Font        = 'Microsoft Sans Serif,10'

$DIB_Logo                        = New-Object system.Windows.Forms.PictureBox
$DIB_Logo.width                  = 274
$DIB_Logo.height                 = 64
$DIB_Logo.location               = New-Object System.Drawing.Point(15,391)
$DIB_Logo.imageLocation          = ".\DidItBetter_logo.png"
$DIB_Logo.SizeMode               = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$UpgradeRMM                      = New-Object system.Windows.Forms.Button
$UpgradeRMM.text                 = "Upgrade Recovery & Migration Manager"
$UpgradeRMM.width                = 250
$UpgradeRMM.height               = 30
$UpgradeRMM.location             = New-Object System.Drawing.Point(15,96)
$UpgradeRMM.Font                 = 'Microsoft Sans Serif,10'

$UpgradeA2E                      = New-Object system.Windows.Forms.Label
$UpgradeA2E.text                 = "Upgrades"
$UpgradeA2E.AutoSize             = $true
$UpgradeA2E.width                = 25
$UpgradeA2E.height               = 10
$UpgradeA2E.location             = New-Object System.Drawing.Point(15,24)
$UpgradeA2E.Font                 = 'Microsoft Sans Serif,14,style=Underline'

$ExchangePermissions             = New-Object system.Windows.Forms.Label
$ExchangePermissions.text        = "Adding Permissions"
$ExchangePermissions.AutoSize    = $true
$ExchangePermissions.width       = 25
$ExchangePermissions.height      = 10
$ExchangePermissions.location    = New-Object System.Drawing.Point(289,24)
$ExchangePermissions.Font        = 'Microsoft Sans Serif,14,style=Underline'

$O365OnPremPermissions           = New-Object system.Windows.Forms.Button
$O365OnPremPermissions.text      = "Office365 & OnPremise Exchange Permissions"
$O365OnPremPermissions.width     = 300
$O365OnPremPermissions.height    = 30
$O365OnPremPermissions.location  = New-Object System.Drawing.Point(289,58)
$O365OnPremPermissions.Font      = 'Microsoft Sans Serif,10'

$Add2OutlookPermissions          = New-Object system.Windows.Forms.Button
$Add2OutlookPermissions.text     = "Add2Outlook Granular Permissions"
$Add2OutlookPermissions.width    = 300
$Add2OutlookPermissions.height   = 30
$Add2OutlookPermissions.location  = New-Object System.Drawing.Point(289,98)
$Add2OutlookPermissions.Font     = 'Microsoft Sans Serif,10'

$TaskCreation                    = New-Object system.Windows.Forms.Button
$TaskCreation.text               = "Automate Permissions on a Schedule"
$TaskCreation.width              = 300
$TaskCreation.height             = 30
$TaskCreation.location           = New-Object System.Drawing.Point(289,138)
$TaskCreation.Font               = 'Microsoft Sans Serif,10'

$Tools                           = New-Object system.Windows.Forms.Label
$Tools.text                      = "Tools"
$Tools.AutoSize                  = $true
$Tools.width                     = 25
$Tools.height                    = 10
$Tools.location                  = New-Object System.Drawing.Point(630,210)
$Tools.Font                      = 'Microsoft Sans Serif,14,style=Underline'

$RunAutoLogon                    = New-Object system.Windows.Forms.Button
$RunAutoLogon.text               = "Auto Logon"
$RunAutoLogon.width              = 147
$RunAutoLogon.height             = 30
$RunAutoLogon.location           = New-Object System.Drawing.Point(629,242)
$RunAutoLogon.Font               = 'Microsoft Sans Serif,10'

$Dir_Sync                        = New-Object system.Windows.Forms.Button
$Dir_Sync.text                   = "Dir Sync"
$Dir_Sync.width                  = 147
$Dir_Sync.height                 = 30
$Dir_Sync.location               = New-Object System.Drawing.Point(629,284)
$Dir_Sync.Font                   = 'Microsoft Sans Serif,10'

$DisableUAC                      = New-Object system.Windows.Forms.Button
$DisableUAC.text                 = "Disable UAC"
$DisableUAC.width                = 147
$DisableUAC.height               = 30
$DisableUAC.location             = New-Object System.Drawing.Point(631,324)
$DisableUAC.Font                 = 'Microsoft Sans Serif,10'

$GroupPolicyResults              = New-Object system.Windows.Forms.Button
$GroupPolicyResults.text         = "Group Policy Results"
$GroupPolicyResults.width        = 147
$GroupPolicyResults.height       = 30
$GroupPolicyResults.location     = New-Object System.Drawing.Point(631,364)
$GroupPolicyResults.Font         = 'Microsoft Sans Serif,10'

$LegacyPowershell                = New-Object system.Windows.Forms.Button
$LegacyPowershell.text           = "Check & Upgrade Powershell"
$LegacyPowershell.width          = 200
$LegacyPowershell.height         = 30
$LegacyPowershell.location       = New-Object System.Drawing.Point(788,242)
$LegacyPowershell.Font           = 'Microsoft Sans Serif,10'

$OutlookAddins                   = New-Object system.Windows.Forms.Button
$OutlookAddins.text              = "Remove Outlook Add-Ins"
$OutlookAddins.width             = 200
$OutlookAddins.height            = 30
$OutlookAddins.location          = New-Object System.Drawing.Point(788,284)
$OutlookAddins.Font              = 'Microsoft Sans Serif,10'

$RegFavs                         = New-Object system.Windows.Forms.Button
$RegFavs.text                    = "Include Registry Favorites"
$RegFavs.width                   = 200
$RegFavs.height                  = 30
$RegFavs.location                = New-Object System.Drawing.Point(788,324)
$RegFavs.Font                    = 'Microsoft Sans Serif,10'

$ExportProfile1                  = New-Object system.Windows.Forms.Button
$ExportProfile1.text             = "Export License & Profile 1"
$ExportProfile1.width            = 200
$ExportProfile1.height           = 30
$ExportProfile1.location         = New-Object System.Drawing.Point(788,364)
$ExportProfile1.Font             = 'Microsoft Sans Serif,10'

$Downloads                       = New-Object system.Windows.Forms.Label
$Downloads.text                  = "Downloads"
$Downloads.AutoSize              = $true
$Downloads.width                 = 25
$Downloads.height                = 10
$Downloads.location              = New-Object System.Drawing.Point(15,210)
$Downloads.Font                  = 'Microsoft Sans Serif,14,style=Underline'

$DownloadLink                    = New-Object system.Windows.Forms.Button
$DownloadLink.text               = "Download Add2Exchange"
$DownloadLink.width              = 180
$DownloadLink.height             = 30
$DownloadLink.location           = New-Object System.Drawing.Point(15,242)
$DownloadLink.Font               = 'Microsoft Sans Serif,10'

$SQLExpress                      = New-Object system.Windows.Forms.Button
$SQLExpress.text                 = "Download SQL Studio"
$SQLExpress.width                = 180
$SQLExpress.height               = 30
$SQLExpress.location             = New-Object System.Drawing.Point(15,282)
$SQLExpress.Font                 = 'Microsoft Sans Serif,10'

$Support                         = New-Object system.Windows.Forms.Label
$Support.text                    = "Get Support"
$Support.AutoSize                = $true
$Support.width                   = 25
$Support.height                  = 10
$Support.location                = New-Object System.Drawing.Point(228,210)
$Support.Font                    = 'Microsoft Sans Serif,14,style=Underline'

$GetSupport                      = New-Object system.Windows.Forms.Button
$GetSupport.text                 = "Need Help? Get Support!"
$GetSupport.width                = 230
$GetSupport.height               = 30
$GetSupport.location             = New-Object System.Drawing.Point(229,242)
$GetSupport.Font                 = 'Microsoft Sans Serif,10'

$SearchDiditbetter               = New-Object system.Windows.Forms.Button
$SearchDiditbetter.text          = "Search DiditBetter"
$SearchDiditbetter.width         = 230
$SearchDiditbetter.height        = 30
$SearchDiditbetter.location      = New-Object System.Drawing.Point(230,282)
$SearchDiditbetter.Font          = 'Microsoft Sans Serif,10'

$GuideA2E                        = New-Object system.Windows.Forms.Button
$GuideA2E.text                   = "Quick Start Guide"
$GuideA2E.width                  = 230
$GuideA2E.height                 = 30
$GuideA2E.location               = New-Object System.Drawing.Point(230,322)
$GuideA2E.Font                   = 'Microsoft Sans Serif,10'

$FTPdownloads                    = New-Object system.Windows.Forms.Button
$FTPdownloads.text               = "FTP Downloads"
$FTPdownloads.width              = 180
$FTPdownloads.height             = 30
$FTPdownloads.location           = New-Object System.Drawing.Point(15,322)
$FTPdownloads.Font               = 'Microsoft Sans Serif,10'

$SyncConcepts                    = New-Object system.Windows.Forms.Label
$SyncConcepts.text               = "Add2Exchange Sync Scenarios How To"
$SyncConcepts.AutoSize           = $true
$SyncConcepts.width              = 25
$SyncConcepts.height             = 10
$SyncConcepts.location           = New-Object System.Drawing.Point(631,24)
$SyncConcepts.Font               = 'Microsoft Sans Serif,14,style=Underline'

$GALSync                         = New-Object system.Windows.Forms.Button
$GALSync.text                    = "GAL Sync"
$GALSync.width                   = 120
$GALSync.height                  = 35
$GALSync.location                = New-Object System.Drawing.Point(631,54)
$GALSync.Font                    = 'Microsoft Sans Serif,10'

$PrivatetoPrivate                = New-Object system.Windows.Forms.Button
$PrivatetoPrivate.text           = "Private to Private"
$PrivatetoPrivate.width          = 120
$PrivatetoPrivate.height         = 35
$PrivatetoPrivate.location       = New-Object System.Drawing.Point(631,94)
$PrivatetoPrivate.Font           = 'Microsoft Sans Serif,10'

$PrivatetoPublic                 = New-Object system.Windows.Forms.Button
$PrivatetoPublic.text            = "Private to Public"
$PrivatetoPublic.width           = 120
$PrivatetoPublic.height          = 35
$PrivatetoPublic.location        = New-Object System.Drawing.Point(768,54)
$PrivatetoPublic.Font            = 'Microsoft Sans Serif,10'

$PublictoPrivate                 = New-Object system.Windows.Forms.Button
$PublictoPrivate.text            = "Public to Private"
$PublictoPrivate.width           = 120
$PublictoPrivate.height          = 35
$PublictoPrivate.location        = New-Object System.Drawing.Point(768,94)
$PublictoPrivate.Font            = 'Microsoft Sans Serif,10'

$PiblictoPublic                  = New-Object system.Windows.Forms.Button
$PiblictoPublic.text             = "Public to Public"
$PiblictoPublic.width            = 120
$PiblictoPublic.height           = 35
$PiblictoPublic.location         = New-Object System.Drawing.Point(631,134)
$PiblictoPublic.Font             = 'Microsoft Sans Serif,10'

$TemplateRels                    = New-Object system.Windows.Forms.Button
$TemplateRels.text               = "Template Creation"
$TemplateRels.width              = 120
$TemplateRels.height             = 35
$TemplateRels.location           = New-Object System.Drawing.Point(768,134)
$TemplateRels.Font               = 'Microsoft Sans Serif,10'

$Migrations                      = New-Object system.Windows.Forms.Button
$Migrations.text                 = "Migrate A2E"
$Migrations.width                = 120
$Migrations.height               = 35
$Migrations.location             = New-Object System.Drawing.Point(905,54)
$Migrations.Font                 = 'Microsoft Sans Serif,10'

$ExchangeMigrate                 = New-Object system.Windows.Forms.Button
$ExchangeMigrate.text            = "Exchange Migration"
$ExchangeMigrate.width           = 120
$ExchangeMigrate.height          = 35
$ExchangeMigrate.location        = New-Object System.Drawing.Point(905,94)
$ExchangeMigrate.Font            = 'Microsoft Sans Serif,10'

$Add2Exchange_Menu.controls.AddRange(@($UpgradeAdd2Exchange,$DIB_Logo,$UpgradeRMM,$UpgradeA2E,$ExchangePermissions,$O365OnPremPermissions,$TaskCreation,$Add2OutlookPermissions,$Tools,$RunAutoLogon,$Dir_Sync,$DisableUAC,$GroupPolicyResults,$LegacyPowershell,$OutlookAddins,$RegFavs,$ExportProfile1,$Downloads,$DownloadLink,$SQLExpress,$Support,$GetSupport,$SearchDiditbetter,$GuideA2E,$FTPdownloads,$SyncConcepts,$GALSync,$PrivatetoPrivate,$PrivatetoPublic,$PublictoPrivate,$PiblictoPublic,$TemplateRels,$Migrations,$ExchangeMigrate))

$UpgradeAdd2Exchange.Add_Click({Start-Process Powershell .\Auto_Upgrade_Add2Exchange.ps1})
$UpgradeRMM.Add_Click({Start-Process Powershell .\Auto_Upgrade_RMM.ps1})
$O365OnPremPermissions.Add_Click({Start-Process Powershell .\PermissionsOnPremOrO365Combined.ps1})
$Add2OutlookPermissions.Add_Click({Start-Process Powershell .\Add2Outlook_Set_Granular_permissions.ps1})
$TaskCreation.Add_Click({Start-Process PowerShell .\Permissions_Task_Creation.ps1})
$RunAutoLogon.Add_Click({Start-Process Powershell .\AutoLogon.exe})
$Dir_Sync.Add_Click({Start-Process Powershell .\Dir_Sync.ps1})
$DisableUAC.Add_Click({Start-Process Powershell .\Disable_UAC.ps1})
$GroupPolicyResults.Add_Click({Start-Process Powershell .\GP_Results.ps1})
$LegacyPowershell.Add_Click({Start-Process Powershell .\Legacy_PowerShell.ps1})
$OutlookAddins.Add_Click({Start-Process Powershell .\Remove_Outlook_Add_ins.ps1})
$RegFavs.Add_Click({ Start-Process Powershell.\Registry_Favorites.ps1})
$ExportProfile1.Add_Click({ Start-Process Powershell.\Export_License_and_Profile1.ps1})
$DownloadLink.Add_Click({Start-Process http://support.DidItBetter.com/Secure/Login.aspx?returnurl=/downloads.aspx})
$SQLExpress.Add_Click({Start-Process ftp://ftp.DidItBetter.com/SQL/SQL2012Management/2012SQLManagementStudio_x64_ENU.exe})
$GetSupport.Add_Click({Start-Process http://support.DidItBetter.com/support-request.aspx})
$SearchDidItBetter.Add_Click({Start-Process http://support.DidItBetter.com/})
$GuideA2E.Add_Click({Start-Process http://guides.diditbetter.com/Add2Exchange_Guide.pdf})
$FTPdownloads.Add_Click({Start-Process ftp.DidItBetter.com})
$GALSync.Add_Click({Start-Process http://guides.diditbetter.com/GAL_Sync_Scenario.pdf})
$PrivatetoPrivate.Add_Click({Start-Process http://guides.diditbetter.com/Private_to_Private_Sync_Scenarios.pdf})
$PrivatetoPublic.Add_Click({Start-Process http://guides.diditbetter.com/Private_to_Public_Sync_Scenarios.pdf})
$PublictoPrivate.Add_Click({Start-Process http://guides.diditbetter.com/Public_to_Private_Sync_Scenarios.pdf})
$PiblictoPublic.Add_Click({Start-Process http://guides.diditbetter.com/Public_to_Public_Sync_Scenarios.pdf})
$TemplateRels.Add_Click({Start-Process http://guides.diditbetter.com/Template_Creation_RGM_Sync_Scenarios.pdf})
$Migrations.Add_Click({Start-Process http://guides.diditbetter.com/Migrating_A2E_Sync_Scenarios.pdf})
$ExchangeMigrate.Add_Click({Start-Process http://guides.diditbetter.com/Migrating_Environments_A2E_Sync_Scenarios.pdf})



#Functions start below here

[void]$Add2Exchange_Menu.ShowDialog()