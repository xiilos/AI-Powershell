<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    A2EMenu
.DESCRIPTION
    Add2Exchange Menu
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Menu               = New-Object system.Windows.Forms.Form
$Add2Exchange_Menu.ClientSize    = '956,461'
$Add2Exchange_Menu.text          = "DiditBetter Software Add2Exchange Setup"
$Add2Exchange_Menu.BackColor     = "#ffffff"
$Add2Exchange_Menu.TopMost       = $false

$Step1Initialize                 = New-Object system.Windows.Forms.Button
$Step1Initialize.text            = "Step 1: Run Me First"
$Step1Initialize.width           = 165
$Step1Initialize.height          = 30
$Step1Initialize.location        = New-Object System.Drawing.Point(14,58)
$Step1Initialize.Font            = 'Microsoft Sans Serif,10'

$DIB_Logo                        = New-Object system.Windows.Forms.PictureBox
$DIB_Logo.width                  = 274
$DIB_Logo.height                 = 64
$DIB_Logo.location               = New-Object System.Drawing.Point(8,346)
$DIB_Logo.imageLocation          = "https://www.diditbetter.com/wp-content/uploads/2017/03/Diditbetter_logo_final.svg"
$DIB_Logo.SizeMode               = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Step2Setup                      = New-Object system.Windows.Forms.Button
$Step2Setup.text                 = "Step 2: 1Click Setup"
$Step2Setup.width                = 165
$Step2Setup.height               = 30
$Step2Setup.location             = New-Object System.Drawing.Point(14,98)
$Step2Setup.Font                 = 'Microsoft Sans Serif,10'

$FirstTimeInstall                = New-Object system.Windows.Forms.Label
$FirstTimeInstall.text           = "First Time Install"
$FirstTimeInstall.AutoSize       = $true
$FirstTimeInstall.width          = 25
$FirstTimeInstall.height         = 10
$FirstTimeInstall.location       = New-Object System.Drawing.Point(14,24)
$FirstTimeInstall.Font           = 'Microsoft Sans Serif,14,style=Underline'

$ExchangePermissions             = New-Object system.Windows.Forms.Label
$ExchangePermissions.text        = "Exchange Permissions"
$ExchangePermissions.AutoSize    = $true
$ExchangePermissions.width       = 25
$ExchangePermissions.height      = 10
$ExchangePermissions.location    = New-Object System.Drawing.Point(227,24)
$ExchangePermissions.Font        = 'Microsoft Sans Serif,14,style=Underline'

$O365OnPremPermissions           = New-Object system.Windows.Forms.Button
$O365OnPremPermissions.text      = "Office365 & OnPremise Permissions"
$O365OnPremPermissions.width     = 230
$O365OnPremPermissions.height    = 30
$O365OnPremPermissions.location  = New-Object System.Drawing.Point(228,58)
$O365OnPremPermissions.Font      = 'Microsoft Sans Serif,10'

$Add2OutlookPermissions          = New-Object system.Windows.Forms.Button
$Add2OutlookPermissions.text     = "A2O Granular Permissions"
$Add2OutlookPermissions.width    = 230
$Add2OutlookPermissions.height   = 30
$Add2OutlookPermissions.location  = New-Object System.Drawing.Point(228,98)
$Add2OutlookPermissions.Font     = 'Microsoft Sans Serif,10'

$Tools                           = New-Object system.Windows.Forms.Label
$Tools.text                      = "Tools"
$Tools.AutoSize                  = $true
$Tools.width                     = 25
$Tools.height                    = 10
$Tools.location                  = New-Object System.Drawing.Point(533,24)
$Tools.Font                      = 'Microsoft Sans Serif,14,style=Underline'

$RunAutoLogon                    = New-Object system.Windows.Forms.Button
$RunAutoLogon.text               = "Auto Logon"
$RunAutoLogon.width              = 147
$RunAutoLogon.height             = 30
$RunAutoLogon.location           = New-Object System.Drawing.Point(533,58)
$RunAutoLogon.Font               = 'Microsoft Sans Serif,10'

$Dir_Sync                        = New-Object system.Windows.Forms.Button
$Dir_Sync.text                   = "Dir Sync"
$Dir_Sync.width                  = 147
$Dir_Sync.height                 = 30
$Dir_Sync.location               = New-Object System.Drawing.Point(533,98)
$Dir_Sync.Font                   = 'Microsoft Sans Serif,10'

$DisableUAC                      = New-Object system.Windows.Forms.Button
$DisableUAC.text                 = "Disable UAC"
$DisableUAC.width                = 147
$DisableUAC.height               = 30
$DisableUAC.location             = New-Object System.Drawing.Point(533,138)
$DisableUAC.Font                 = 'Microsoft Sans Serif,10'

$GroupPolicyResults              = New-Object system.Windows.Forms.Button
$GroupPolicyResults.text         = "Group Policy Results"
$GroupPolicyResults.width        = 147
$GroupPolicyResults.height       = 30
$GroupPolicyResults.location     = New-Object System.Drawing.Point(533,178)
$GroupPolicyResults.Font         = 'Microsoft Sans Serif,10'

$LegacyPowershell                = New-Object system.Windows.Forms.Button
$LegacyPowershell.text           = "Check & Upgrade Powershell"
$LegacyPowershell.width          = 200
$LegacyPowershell.height         = 30
$LegacyPowershell.location       = New-Object System.Drawing.Point(718,58)
$LegacyPowershell.Font           = 'Microsoft Sans Serif,10'

$OutlookAddins                   = New-Object system.Windows.Forms.Button
$OutlookAddins.text              = "Remove Outlook Add-Ins"
$OutlookAddins.width             = 200
$OutlookAddins.height            = 30
$OutlookAddins.location          = New-Object System.Drawing.Point(718,98)
$OutlookAddins.Font              = 'Microsoft Sans Serif,10'

$RegFavs                         = New-Object system.Windows.Forms.Button
$RegFavs.text                    = "Include Registry Favorites"
$RegFavs.width                   = 200
$RegFavs.height                  = 30
$RegFavs.location                = New-Object System.Drawing.Point(718,138)
$RegFavs.Font                    = 'Microsoft Sans Serif,10'

$ExportProfile1                  = New-Object system.Windows.Forms.Button
$ExportProfile1.text             = "Export License & Profile 1"
$ExportProfile1.width            = 200
$ExportProfile1.height           = 30
$ExportProfile1.location         = New-Object System.Drawing.Point(718,178)
$ExportProfile1.Font             = 'Microsoft Sans Serif,10'

$Downloads                       = New-Object system.Windows.Forms.Label
$Downloads.text                  = "Downloads"
$Downloads.AutoSize              = $true
$Downloads.width                 = 25
$Downloads.height                = 10
$Downloads.location              = New-Object System.Drawing.Point(14,166)
$Downloads.Font                  = 'Microsoft Sans Serif,14,style=Underline'

$DownloadLink                    = New-Object system.Windows.Forms.Button
$DownloadLink.text               = "Download Add2Exchange"
$DownloadLink.width              = 180
$DownloadLink.height             = 30
$DownloadLink.location           = New-Object System.Drawing.Point(14,198)
$DownloadLink.Font               = 'Microsoft Sans Serif,10'

$SQLExpress                      = New-Object system.Windows.Forms.Button
$SQLExpress.text                 = "Download SQL Studio"
$SQLExpress.width                = 180
$SQLExpress.height               = 30
$SQLExpress.location             = New-Object System.Drawing.Point(14,238)
$SQLExpress.Font                 = 'Microsoft Sans Serif,10'

$Support                         = New-Object system.Windows.Forms.Label
$Support.text                    = "Get Support"
$Support.AutoSize                = $true
$Support.width                   = 25
$Support.height                  = 10
$Support.location                = New-Object System.Drawing.Point(227,166)
$Support.Font                    = 'Microsoft Sans Serif,14,style=Underline'

$GetSupport                      = New-Object system.Windows.Forms.Button
$GetSupport.text                 = "Need Help? Get Support!"
$GetSupport.width                = 230
$GetSupport.height               = 30
$GetSupport.location             = New-Object System.Drawing.Point(228,198)
$GetSupport.Font                 = 'Microsoft Sans Serif,10'

$SearchDiditbetter               = New-Object system.Windows.Forms.Button
$SearchDiditbetter.text          = "Search DiditBetter"
$SearchDiditbetter.width         = 230
$SearchDiditbetter.height        = 30
$SearchDiditbetter.location      = New-Object System.Drawing.Point(229,238)
$SearchDiditbetter.Font          = 'Microsoft Sans Serif,10'

$GuideA2E                        = New-Object system.Windows.Forms.Button
$GuideA2E.text                   = "Quick Start Guide"
$GuideA2E.width                  = 230
$GuideA2E.height                 = 30
$GuideA2E.location               = New-Object System.Drawing.Point(229,278)
$GuideA2E.Font                   = 'Microsoft Sans Serif,10'

$FTPdownloads                    = New-Object system.Windows.Forms.Button
$FTPdownloads.text               = "FTP Downloads"
$FTPdownloads.width              = 180
$FTPdownloads.height             = 30
$FTPdownloads.location           = New-Object System.Drawing.Point(14,278)
$FTPdownloads.Font               = 'Microsoft Sans Serif,10'

$SyncConcepts                    = New-Object system.Windows.Forms.Label
$SyncConcepts.text               = "Add2Exchange Sync Scenarios How To"
$SyncConcepts.AutoSize           = $true
$SyncConcepts.width              = 25
$SyncConcepts.height             = 10
$SyncConcepts.location           = New-Object System.Drawing.Point(533,235)
$SyncConcepts.Font               = 'Microsoft Sans Serif,14,style=Underline'

$GALSync                         = New-Object system.Windows.Forms.Button
$GALSync.text                    = "GAL Sync"
$GALSync.width                   = 120
$GALSync.height                  = 30
$GALSync.location                = New-Object System.Drawing.Point(533,265)
$GALSync.Font                    = 'Microsoft Sans Serif,10'

$PrivatetoProvate                = New-Object system.Windows.Forms.Button
$PrivatetoProvate.text           = "Private to Private"
$PrivatetoProvate.width          = 120
$PrivatetoProvate.height         = 30
$PrivatetoProvate.location       = New-Object System.Drawing.Point(533,305)
$PrivatetoProvate.Font           = 'Microsoft Sans Serif,10'

$PrivatetoPublic                 = New-Object system.Windows.Forms.Button
$PrivatetoPublic.text            = "Private to Public"
$PrivatetoPublic.width           = 120
$PrivatetoPublic.height          = 30
$PrivatetoPublic.location        = New-Object System.Drawing.Point(670,265)
$PrivatetoPublic.Font            = 'Microsoft Sans Serif,10'

$PublictoPrivate                 = New-Object system.Windows.Forms.Button
$PublictoPrivate.text            = "Public to Private"
$PublictoPrivate.width           = 120
$PublictoPrivate.height          = 30
$PublictoPrivate.location        = New-Object System.Drawing.Point(670,305)
$PublictoPrivate.Font            = 'Microsoft Sans Serif,10'

$PiblictoPublic                  = New-Object system.Windows.Forms.Button
$PiblictoPublic.text             = "Public to Public"
$PiblictoPublic.width            = 120
$PiblictoPublic.height           = 30
$PiblictoPublic.location         = New-Object System.Drawing.Point(533,345)
$PiblictoPublic.Font             = 'Microsoft Sans Serif,10'

$TemplateRels                    = New-Object system.Windows.Forms.Button
$TemplateRels.text               = "Template Creation"
$TemplateRels.width              = 120
$TemplateRels.height             = 30
$TemplateRels.location           = New-Object System.Drawing.Point(670,345)
$TemplateRels.Font               = 'Microsoft Sans Serif,10'

$Migrations                      = New-Object system.Windows.Forms.Button
$Migrations.text                 = "Migrate A2E"
$Migrations.width                = 120
$Migrations.height               = 30
$Migrations.location             = New-Object System.Drawing.Point(807,265)
$Migrations.Font                 = 'Microsoft Sans Serif,10'

$ExchangeMigrate                 = New-Object system.Windows.Forms.Button
$ExchangeMigrate.text            = "Exchange Migration"
$ExchangeMigrate.width           = 120
$ExchangeMigrate.height          = 30
$ExchangeMigrate.location        = New-Object System.Drawing.Point(807,305)
$ExchangeMigrate.Font            = 'Microsoft Sans Serif,10'

$Add2Exchange_Menu.controls.AddRange(@($Step1Initialize,$DIB_Logo,$Step2Setup,$FirstTimeInstall,$ExchangePermissions,$O365OnPremPermissions,$Add2OutlookPermissions,$Tools,$RunAutoLogon,$Dir_Sync,$DisableUAC,$GroupPolicyResults,$LegacyPowershell,$OutlookAddins,$RegFavs,$ExportProfile1,$Downloads,$DownloadLink,$SQLExpress,$Support,$GetSupport,$SearchDiditbetter,$GuideA2E,$FTPdownloads,$SyncConcepts,$GALSync,$PrivatetoProvate,$PrivatetoPublic,$PublictoPrivate,$PiblictoPublic,$TemplateRels,$Migrations,$ExchangeMigrate))

$Step1Initialize.Add_Click({  })
$Step2Setup.Add_Click({  })
$O365OnPremPermissions.Add_Click({  })
$Add2OutlookPermissions.Add_Click({  })
$RunAutoLogon.Add_Click({  })
$Dir_Sync.Add_Click({  })
$DisableUAC.Add_Click({  })
$GroupPolicyResults.Add_Click({  })
$LegacyPowershell.Add_Click({  })
$OutlookAddins.Add_Click({  })
$RegFavs.Add_Click({  })
$ExportProfile1.Add_Click({  })
$DownloadLink.Add_Click({  })
$SQLExpress.Add_Click({  })
$GetSupport.Add_Click({  })
$SearchDiditbetter.Add_Click({  })
$GuideA2E.Add_Click({  })
$GALSync.Add_Click({  })
$PrivatetoProvate.Add_Click({  })
$PrivatetoPublic.Add_Click({  })
$PublictoPrivate.Add_Click({  })
$PiblictoPublic.Add_Click({  })
$FTPdownloads.Add_Click({  })
$TemplateRels.Add_Click({  })
$Migrations.Add_Click({  })
$ExchangeMigrate.Add_Click({  })



#Write your logic code here

[void]$Add2Exchange_Menu.ShowDialog()