<#
.NAME
    A2E Menu
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$DidItBetterSupportMenu          = New-Object system.Windows.Forms.Form
$DidItBetterSupportMenu.ClientSize  = New-Object System.Drawing.Point(542,725)
$DidItBetterSupportMenu.text     = "DidItBetter Software Support Menu"
$DidItBetterSupportMenu.TopMost  = $false
$DidItBetterSupportMenu.BackColor  = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$Upgrades                        = New-Object system.Windows.Forms.Label
$Upgrades.text                   = "Upgrades"
$Upgrades.AutoSize               = $true
$Upgrades.width                  = 150
$Upgrades.height                 = 10
$Upgrades.location               = New-Object System.Drawing.Point(255,10)
$Upgrades.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Add2ExchangeUpgrade             = New-Object system.Windows.Forms.Label
$Add2ExchangeUpgrade.text        = "Upgrade Add2Exchange Enterprise"
$Add2ExchangeUpgrade.AutoSize    = $true
$Add2ExchangeUpgrade.width       = 150
$Add2ExchangeUpgrade.height      = 10
$Add2ExchangeUpgrade.location    = New-Object System.Drawing.Point(255,30)
$Add2ExchangeUpgrade.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$UpgradeRMM                      = New-Object system.Windows.Forms.Label
$UpgradeRMM.text                 = "Upgrade Recovery and Migration Manager"
$UpgradeRMM.AutoSize             = $true
$UpgradeRMM.width                = 150
$UpgradeRMM.height               = 10
$UpgradeRMM.location             = New-Object System.Drawing.Point(255,45)
$UpgradeRMM.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Permissions                     = New-Object system.Windows.Forms.Label
$Permissions.text                = "Permissions"
$Permissions.AutoSize            = $true
$Permissions.width               = 25
$Permissions.height              = 10
$Permissions.location            = New-Object System.Drawing.Point(255,140)
$Permissions.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$O365ExchangePermissions         = New-Object system.Windows.Forms.Label
$O365ExchangePermissions.text    = "Office 365 and On-Premise Exchange Permissions"
$O365ExchangePermissions.AutoSize  = $true
$O365ExchangePermissions.width   = 150
$O365ExchangePermissions.height  = 10
$O365ExchangePermissions.location  = New-Object System.Drawing.Point(255,160)
$O365ExchangePermissions.Font    = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$A2OPermissions                  = New-Object system.Windows.Forms.Label
$A2OPermissions.text             = "Add2Outlook Granular Permissions"
$A2OPermissions.AutoSize         = $true
$A2OPermissions.width            = 150
$A2OPermissions.height           = 10
$A2OPermissions.location         = New-Object System.Drawing.Point(255,205)
$A2OPermissions.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$AutoPermissions                 = New-Object system.Windows.Forms.Label
$AutoPermissions.text            = "Automate Permissions on a Schedule"
$AutoPermissions.AutoSize        = $true
$AutoPermissions.width           = 150
$AutoPermissions.height          = 10
$AutoPermissions.location        = New-Object System.Drawing.Point(255,175)
$AutoPermissions.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Downloads                       = New-Object system.Windows.Forms.Label
$Downloads.text                  = "Downloads"
$Downloads.AutoSize              = $true
$Downloads.width                 = 150
$Downloads.height                = 10
$Downloads.location              = New-Object System.Drawing.Point(15,363)
$Downloads.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$DownloadAdd2Exchange            = New-Object system.Windows.Forms.Label
$DownloadAdd2Exchange.text       = "Download Add2Exchange"
$DownloadAdd2Exchange.AutoSize   = $true
$DownloadAdd2Exchange.width      = 150
$DownloadAdd2Exchange.height     = 10
$DownloadAdd2Exchange.location   = New-Object System.Drawing.Point(15,383)
$DownloadAdd2Exchange.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DownloadToolKit                 = New-Object system.Windows.Forms.Label
$DownloadToolKit.text            = "Download ToolKit"
$DownloadToolKit.AutoSize        = $true
$DownloadToolKit.width           = 150
$DownloadToolKit.height          = 10
$DownloadToolKit.location        = New-Object System.Drawing.Point(15,398)
$DownloadToolKit.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DownloadSQL                     = New-Object system.Windows.Forms.Label
$DownloadSQL.text                = "Download SQL Studio"
$DownloadSQL.AutoSize            = $true
$DownloadSQL.width               = 150
$DownloadSQL.height              = 10
$DownloadSQL.location            = New-Object System.Drawing.Point(15,413)
$DownloadSQL.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$FTPDownloads                    = New-Object system.Windows.Forms.Label
$FTPDownloads.text               = "FTP Downloads"
$FTPDownloads.AutoSize           = $true
$FTPDownloads.width              = 150
$FTPDownloads.height             = 10
$FTPDownloads.location           = New-Object System.Drawing.Point(15,443)
$FTPDownloads.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$GetSupport                      = New-Object system.Windows.Forms.Label
$GetSupport.text                 = "Get Support"
$GetSupport.AutoSize             = $true
$GetSupport.width                = 150
$GetSupport.height               = 10
$GetSupport.location             = New-Object System.Drawing.Point(15,10)
$GetSupport.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$GetHelp                         = New-Object system.Windows.Forms.Label
$GetHelp.text                    = "Need Help? Open a Ticket!"
$GetHelp.AutoSize                = $true
$GetHelp.width                   = 150
$GetHelp.height                  = 10
$GetHelp.location                = New-Object System.Drawing.Point(15,30)
$GetHelp.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$SearchDidItBetter               = New-Object system.Windows.Forms.Label
$SearchDidItBetter.text          = "Search DidItBetter"
$SearchDidItBetter.AutoSize      = $true
$SearchDidItBetter.width         = 150
$SearchDidItBetter.height        = 10
$SearchDidItBetter.location      = New-Object System.Drawing.Point(15,60)
$SearchDidItBetter.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$QuickStartGuide                 = New-Object system.Windows.Forms.Label
$QuickStartGuide.text            = "Quick Start Guide"
$QuickStartGuide.AutoSize        = $true
$QuickStartGuide.width           = 150
$QuickStartGuide.height          = 10
$QuickStartGuide.location        = New-Object System.Drawing.Point(15,45)
$QuickStartGuide.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Tools                           = New-Object system.Windows.Forms.Label
$Tools.text                      = "Tools"
$Tools.AutoSize                  = $true
$Tools.width                     = 150
$Tools.height                    = 10
$Tools.location                  = New-Object System.Drawing.Point(255,363)
$Tools.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$AutoLogon                       = New-Object system.Windows.Forms.Label
$AutoLogon.text                  = "Auto Logon Encrypted"
$AutoLogon.AutoSize              = $true
$AutoLogon.width                 = 150
$AutoLogon.height                = 10
$AutoLogon.location              = New-Object System.Drawing.Point(255,220)
$AutoLogon.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DirSync                         = New-Object system.Windows.Forms.Label
$DirSync.text                    = "Azure Directory Sync"
$DirSync.AutoSize                = $true
$DirSync.width                   = 150
$DirSync.height                  = 10
$DirSync.location                = New-Object System.Drawing.Point(255,428)
$DirSync.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DisableUAC                      = New-Object system.Windows.Forms.Label
$DisableUAC.text                 = "Test and Disable User Account Control"
$DisableUAC.AutoSize             = $true
$DisableUAC.width                = 150
$DisableUAC.height               = 10
$DisableUAC.location             = New-Object System.Drawing.Point(255,558)
$DisableUAC.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$GroupPolicyResults              = New-Object system.Windows.Forms.Label
$GroupPolicyResults.text         = "Group Policy Results"
$GroupPolicyResults.AutoSize     = $true
$GroupPolicyResults.width        = 150
$GroupPolicyResults.height       = 10
$GroupPolicyResults.location     = New-Object System.Drawing.Point(255,398)
$GroupPolicyResults.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$CheckPowerShell                 = New-Object system.Windows.Forms.Label
$CheckPowerShell.text            = "Check and Upgrade PowerShell"
$CheckPowerShell.AutoSize        = $true
$CheckPowerShell.width           = 150
$CheckPowerShell.height          = 10
$CheckPowerShell.location        = New-Object System.Drawing.Point(255,573)
$CheckPowerShell.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DisableOSC                      = New-Object system.Windows.Forms.Label
$DisableOSC.text                 = "Required-Disable Outlook Social Connector"
$DisableOSC.AutoSize             = $true
$DisableOSC.width                = 150
$DisableOSC.height               = 10
$DisableOSC.location             = New-Object System.Drawing.Point(255,285)
$DisableOSC.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$AddRegistryFavorites            = New-Object system.Windows.Forms.Label
$AddRegistryFavorites.text       = "Add Registry Favorites"
$AddRegistryFavorites.AutoSize   = $true
$AddRegistryFavorites.width      = 150
$AddRegistryFavorites.height     = 10
$AddRegistryFavorites.location   = New-Object System.Drawing.Point(255,235)
$AddRegistryFavorites.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Reset_A2E_Passwords             = New-Object system.Windows.Forms.Label
$Reset_A2E_Passwords.text        = "Reset The Add2Exchange password"
$Reset_A2E_Passwords.AutoSize    = $true
$Reset_A2E_Passwords.width       = 150
$Reset_A2E_Passwords.height      = 10
$Reset_A2E_Passwords.location    = New-Object System.Drawing.Point(255,190)
$Reset_A2E_Passwords.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$SyncScenarios                   = New-Object system.Windows.Forms.Label
$SyncScenarios.text              = "Add2Exchange Sync Scenarios"
$SyncScenarios.AutoSize          = $true
$SyncScenarios.width             = 150
$SyncScenarios.height            = 10
$SyncScenarios.location          = New-Object System.Drawing.Point(15,140)
$SyncScenarios.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$GALSync                         = New-Object system.Windows.Forms.Label
$GALSync.text                    = "GAL Synchronization"
$GALSync.AutoSize                = $true
$GALSync.width                   = 150
$GALSync.height                  = 10
$GALSync.location                = New-Object System.Drawing.Point(15,160)
$GALSync.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$PrivatetoPrivate                = New-Object system.Windows.Forms.Label
$PrivatetoPrivate.text           = "Private to Private Relationship"
$PrivatetoPrivate.AutoSize       = $true
$PrivatetoPrivate.width          = 150
$PrivatetoPrivate.height         = 10
$PrivatetoPrivate.location       = New-Object System.Drawing.Point(15,175)
$PrivatetoPrivate.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$PublictoPublic                  = New-Object system.Windows.Forms.Label
$PublictoPublic.text             = "Public to Public Relationship"
$PublictoPublic.AutoSize         = $true
$PublictoPublic.width            = 150
$PublictoPublic.height           = 10
$PublictoPublic.location         = New-Object System.Drawing.Point(15,190)
$PublictoPublic.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$PrivatetoPublic                 = New-Object system.Windows.Forms.Label
$PrivatetoPublic.text            = "Private to Public Relationship"
$PrivatetoPublic.AutoSize        = $true
$PrivatetoPublic.width           = 150
$PrivatetoPublic.height          = 10
$PrivatetoPublic.location        = New-Object System.Drawing.Point(15,205)
$PrivatetoPublic.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$PublictoPrivate                 = New-Object system.Windows.Forms.Label
$PublictoPrivate.text            = "Public to Private Relationship"
$PublictoPrivate.AutoSize        = $true
$PublictoPrivate.width           = 150
$PublictoPrivate.height          = 10
$PublictoPrivate.location        = New-Object System.Drawing.Point(15,220)
$PublictoPrivate.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$TemplateCreation                = New-Object system.Windows.Forms.Label
$TemplateCreation.text           = "Template Creation"
$TemplateCreation.AutoSize       = $true
$TemplateCreation.width          = 150
$TemplateCreation.height         = 10
$TemplateCreation.location       = New-Object System.Drawing.Point(15,235)
$TemplateCreation.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$MigrateA2E                      = New-Object system.Windows.Forms.Label
$MigrateA2E.text                 = "Migrate Add2Exchange to a New Box"
$MigrateA2E.AutoSize             = $true
$MigrateA2E.width                = 150
$MigrateA2E.height               = 10
$MigrateA2E.location             = New-Object System.Drawing.Point(15,250)
$MigrateA2E.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$ExhangeMigration                = New-Object system.Windows.Forms.Label
$ExhangeMigration.text           = "Exchange Migration with Add2Exchange"
$ExhangeMigration.AutoSize       = $true
$ExhangeMigration.width          = 150
$ExhangeMigration.height         = 10
$ExhangeMigration.location       = New-Object System.Drawing.Point(15,265)
$ExhangeMigration.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$UpgradeToolkit                  = New-Object system.Windows.Forms.Label
$UpgradeToolkit.text             = "Upgrade Add2Outlook ToolKit"
$UpgradeToolkit.AutoSize         = $true
$UpgradeToolkit.width            = 150
$UpgradeToolkit.height           = 10
$UpgradeToolkit.location         = New-Object System.Drawing.Point(255,60)
$UpgradeToolkit.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$A2EDiags                        = New-Object system.Windows.Forms.Label
$A2EDiags.text                   = "Get A2E Diags"
$A2EDiags.AutoSize               = $true
$A2EDiags.width                  = 150
$A2EDiags.height                 = 10
$A2EDiags.location               = New-Object System.Drawing.Point(15,428)
$A2EDiags.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$CreateSupporttext               = New-Object system.Windows.Forms.Label
$CreateSupporttext.text          = "Create Support Setup Details"
$CreateSupporttext.AutoSize      = $true
$CreateSupporttext.width         = 150
$CreateSupporttext.height        = 10
$CreateSupporttext.location      = New-Object System.Drawing.Point(255,413)
$CreateSupporttext.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Revision                        = New-Object system.Windows.Forms.Label
$Revision.text                   = "Rev. 4.27.7"
$Revision.AutoSize               = $true
$Revision.width                  = 25
$Revision.height                 = 10
$Revision.location               = New-Object System.Drawing.Point(455,700)
$Revision.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',8,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

$UpgradeAdd2Outlook              = New-Object system.Windows.Forms.Label
$UpgradeAdd2Outlook.text         = "Upgrade Add2Outlook"
$UpgradeAdd2Outlook.AutoSize     = $true
$UpgradeAdd2Outlook.width        = 25
$UpgradeAdd2Outlook.height       = 10
$UpgradeAdd2Outlook.location     = New-Object System.Drawing.Point(255,75)
$UpgradeAdd2Outlook.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$ToolTip1                        = New-Object system.Windows.Forms.ToolTip
$ToolTip1.ToolTipTitle           = "Help"
$ToolTip1.isBalloon              = $true

$AD_Photos                       = New-Object system.Windows.Forms.Label
$AD_Photos.text                  = "Export Active Directory Photos"
$AD_Photos.AutoSize              = $true
$AD_Photos.width                 = 150
$AD_Photos.height                = 10
$AD_Photos.location              = New-Object System.Drawing.Point(255,300)
$AD_Photos.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$MSExchangeDelegate              = New-Object system.Windows.Forms.Label
$MSExchangeDelegate.text         = "MSExchange Delegate Automapping Fix"
$MSExchangeDelegate.AutoSize     = $true
$MSExchangeDelegate.width        = 25
$MSExchangeDelegate.height       = 10
$MSExchangeDelegate.location     = New-Object System.Drawing.Point(255,543)
$MSExchangeDelegate.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$ExchangeShell                   = New-Object system.Windows.Forms.Label
$ExchangeShell.text              = "Shell Into Exchange or Office 365"
$ExchangeShell.AutoSize          = $true
$ExchangeShell.width             = 150
$ExchangeShell.height            = 10
$ExchangeShell.location          = New-Object System.Drawing.Point(255,383)
$ExchangeShell.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$CommandsList                    = New-Object system.Windows.Forms.Label
$CommandsList.text               = "Show Me Commands"
$CommandsList.AutoSize           = $true
$CommandsList.width              = 150
$CommandsList.height             = 10
$CommandsList.location           = New-Object System.Drawing.Point(15,75)
$CommandsList.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$A2E_Migration_Wizard            = New-Object system.Windows.Forms.Label
$A2E_Migration_Wizard.text       = "Add2Exchange Migration Wizard"
$A2E_Migration_Wizard.AutoSize   = $true
$A2E_Migration_Wizard.width      = 150
$A2E_Migration_Wizard.height     = 10
$A2E_Migration_Wizard.location   = New-Object System.Drawing.Point(255,588)
$A2E_Migration_Wizard.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Configuration                   = New-Object system.Windows.Forms.Label
$Configuration.text              = "Configuration"
$Configuration.AutoSize          = $true
$Configuration.width             = 25
$Configuration.height            = 10
$Configuration.location          = New-Object System.Drawing.Point(255,265)
$Configuration.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$FixesEvents                     = New-Object system.Windows.Forms.Label
$FixesEvents.text                = "Fixes and Events"
$FixesEvents.AutoSize            = $true
$FixesEvents.width               = 25
$FixesEvents.height              = 10
$FixesEvents.location            = New-Object System.Drawing.Point(255,508)
$FixesEvents.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$ModernAuth                      = New-Object system.Windows.Forms.Label
$ModernAuth.text                 = "Disable Modern Authentication"
$ModernAuth.AutoSize             = $true
$ModernAuth.width                = 25
$ModernAuth.height               = 10
$ModernAuth.location             = New-Object System.Drawing.Point(255,528)
$ModernAuth.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Shortcuts                       = New-Object system.Windows.Forms.Label
$Shortcuts.text                  = "Shortcuts"
$Shortcuts.AutoSize              = $true
$Shortcuts.width                 = 25
$Shortcuts.height                = 10
$Shortcuts.location              = New-Object System.Drawing.Point(15,508)
$Shortcuts.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$DIBMMC                          = New-Object system.Windows.Forms.Label
$DIBMMC.text                     = "DidItBetter MMC"
$DIBMMC.AutoSize                 = $true
$DIBMMC.width                    = 25
$DIBMMC.height                   = 10
$DIBMMC.location                 = New-Object System.Drawing.Point(15,528)
$DIBMMC.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$RegistryEditor                  = New-Object system.Windows.Forms.Label
$RegistryEditor.text             = "Registry Editor"
$RegistryEditor.AutoSize         = $true
$RegistryEditor.width            = 25
$RegistryEditor.height           = 10
$RegistryEditor.location         = New-Object System.Drawing.Point(15,543)
$RegistryEditor.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$DIBDirectory                    = New-Object system.Windows.Forms.Label
$DIBDirectory.text               = "DidItBetter Directory"
$DIBDirectory.AutoSize           = $true
$DIBDirectory.width              = 25
$DIBDirectory.height             = 10
$DIBDirectory.location           = New-Object System.Drawing.Point(15,558)
$DIBDirectory.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$TaskScheduler                   = New-Object system.Windows.Forms.Label
$TaskScheduler.text              = "Task Scheduler"
$TaskScheduler.AutoSize          = $true
$TaskScheduler.width             = 25
$TaskScheduler.height            = 10
$TaskScheduler.location          = New-Object System.Drawing.Point(15,573)
$TaskScheduler.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$WindowsDefender                 = New-Object system.Windows.Forms.Label
$WindowsDefender.text            = "Windows Defender Exclusions"
$WindowsDefender.AutoSize        = $true
$WindowsDefender.width           = 150
$WindowsDefender.height          = 10
$WindowsDefender.location        = New-Object System.Drawing.Point(255,315)
$WindowsDefender.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$Outlook_Tools_Menu              = New-Object system.Windows.Forms.Label
$Outlook_Tools_Menu.text         = "Outlook Tools"
$Outlook_Tools_Menu.AutoSize     = $true
$Outlook_Tools_Menu.width        = 150
$Outlook_Tools_Menu.height       = 10
$Outlook_Tools_Menu.location     = New-Object System.Drawing.Point(255,443)
$Outlook_Tools_Menu.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$A2ESQLBackUp                    = New-Object system.Windows.Forms.Label
$A2ESQLBackUp.text               = "Backup Add2Exchange SQL DB"
$A2ESQLBackUp.AutoSize           = $true
$A2ESQLBackUp.width              = 150
$A2ESQLBackUp.height             = 10
$A2ESQLBackUp.location           = New-Object System.Drawing.Point(255,458)
$A2ESQLBackUp.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',9)

$SQLUpgrade                      = New-Object system.Windows.Forms.Label
$SQLUpgrade.text                 = "Upgrade SQL 2008 to 2012"
$SQLUpgrade.AutoSize             = $true
$SQLUpgrade.width                = 25
$SQLUpgrade.height               = 10
$SQLUpgrade.location             = New-Object System.Drawing.Point(255,473)
$SQLUpgrade.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',9)


$ToolTip1.SetToolTip($Add2ExchangeUpgrade,'This will Upgrade your Current Version of Add2Exchange Enterprise to the Latest Version')
$ToolTip1.SetToolTip($UpgradeRMM,'This will Upgrade your Current Version of Recovery and Migration Manager to the Latest Version')
$ToolTip1.SetToolTip($O365ExchangePermissions,'Run this to Add Permissions to any users that will be syncing with Add2Exchange')
$ToolTip1.SetToolTip($A2OPermissions,'Run this to Add Granular Folder Permissions to Users you will be syncing with Add2Outlook')
$ToolTip1.SetToolTip($AutoPermissions,'Create a Windows Task to add Permissions to users on a Schedule')
$ToolTip1.SetToolTip($AutoLogon,'Run the Microsoft Sys Internals AutoLogon GUI')
$ToolTip1.SetToolTip($DirSync,'PowerShell Script to run an on Demand Directory Sync')
$ToolTip1.SetToolTip($DisableUAC,'PowerShell Script to Disable User Account Control from within the Registry')
$ToolTip1.SetToolTip($GroupPolicyResults,'PowerShell Script to check current Group Policy on this Appliance ')
$ToolTip1.SetToolTip($CheckPowerShell,'PowerShell Script to check and Upgrade PowerShell if needed')
$ToolTip1.SetToolTip($DisableOSC,'PowerShell Script to Unload Outlook Social Connector')
$ToolTip1.SetToolTip($AddRegistryFavorites,'PowerShell Script to Add Add2Exchange Favorites in the Registry')
$ToolTip1.SetToolTip($Reset_A2E_Passwords,'PowerShell Script to Reset the Add2Exchange Password in Both Locations')
$ToolTip1.SetToolTip($GALSync,'Click for more information on How To Setup a Global Address Sync ')
$ToolTip1.SetToolTip($PrivatetoPrivate,'Click for more information on How To Setup a Private to Private Relationship')
$ToolTip1.SetToolTip($PublictoPublic,'Click for more information on How To Setup a Public to Public Relationship')
$ToolTip1.SetToolTip($PrivatetoPublic,'Click for more information on How To Setup a Private to Public')
$ToolTip1.SetToolTip($PublictoPrivate,'Click for more information on How To Setup a Public to Private')
$ToolTip1.SetToolTip($TemplateCreation,'Click for more information on How To Setup Templates for Synchronization with Distribution Groups')
$ToolTip1.SetToolTip($MigrateA2E,'Click for more information on How To Migrate Add2Exchange onto a new Appliance')
$ToolTip1.SetToolTip($ExhangeMigration,'Click for more information on How To Setup Add2Exchange before or After an Exchange or Office 365 Migration')
$ToolTip1.SetToolTip($UpgradeToolkit,'This will Upgrade your Current Version of Add2Outlook Toolkit to the Latest Version')
$ToolTip1.SetToolTip($CreateSupporttext,'PowerShell Script to Create a Support text detailing this Installation')
$ToolTip1.SetToolTip($UpgradeAdd2Outlook,'This will Upgrade your Current Version of Add2Outlook to the Latest Version')
$ToolTip1.SetToolTip($AD_Photos,'This tool will Export User Active Directory Photos')
$ToolTip1.SetToolTip($MSExchangeDelegate,'Removes the List Link in AD for Users that are getting Auto Mapped in Outlook')
$ToolTip1.SetToolTip($ExchangeShell,'Log Into Exchange or Office 365 Shell')
$ToolTip1.SetToolTip($CommandsList,'Show the List of Commands to Run for Add2Exchange Permissions')
$ToolTip1.SetToolTip($A2E_Migration_Wizard,'Wizard for Migrating Add2Exchange to another Appliance')
$ToolTip1.SetToolTip($ModernAuth,'Disables Modern Authentication for Outlook 2016 and Up')
$ToolTip1.SetToolTip($DIBMMC,'Opens the DidItBetter Event Viewer')
$ToolTip1.SetToolTip($RegistryEditor,'Open the registry editor')
$ToolTip1.SetToolTip($DIBDirectory,'Opens the DIB Directory')
$ToolTip1.SetToolTip($TaskScheduler,'Opens Task Scheduler')
$ToolTip1.SetToolTip($WindowsDefender,'This tool will add the proper Exclusions to Windows Defender')
$ToolTip1.SetToolTip($A2ESQLBackUp,'Backup the Add2Exchange SQL Database')
$ToolTip1.SetToolTip($SQLUpgrade,'Upgrade SQL Database from 2008-2008R2 to 2012')

$DidItBetterSupportMenu.controls.AddRange(@($Upgrades,$Add2ExchangeUpgrade,$UpgradeRMM,$Permissions,$O365ExchangePermissions,$A2OPermissions,$AutoPermissions,$Downloads,$DownloadAdd2Exchange,
$DownloadToolKit,$DownloadSQL,$FTPDownloads,$GetSupport,$GetHelp,$SearchDidItBetter,$QuickStartGuide,$Tools,$AutoLogon,$DirSync,$DisableUAC,$GroupPolicyResults,$CheckPowerShell,$DisableOSC,
$AddRegistryFavorites,$Reset_A2E_Passwords,$SyncScenarios,$GALSync,$PrivatetoPrivate,$PublictoPublic,$PrivatetoPublic,$PublictoPrivate,$TemplateCreation,$MigrateA2E,$ExhangeMigration,$DidItBetterLogo,
$UpgradeToolkit,$A2EDiags,$CreateSupporttext,$Revision,$UpgradeAdd2Outlook,$AD_Photos,$MSExchangeDelegate,$ExchangeShell,$CommandsList,$A2E_Migration_Wizard,$Configuration,$FixesEvents,$ModernAuth,
$Shortcuts,$DIBMMC,$RegistryEditor,$DIBDirectory,$TaskScheduler,$WindowsDefender,$Outlook_Tools_Menu,$A2ESQLBackUp,$SQLUpgrade))

$Add2ExchangeUpgrade.Add_Click( { Start-Process Powershell .\Auto_Upgrade_Add2Exchange.ps1 })
$UpgradeRMM.Add_Click( { Start-Process Powershell .\Auto_Upgrade_RMM.ps1 })
$UpgradeToolkit.Add_Click( { Start-Process Powershell .\Auto_Upgrade_ToolKit.ps1 })
$UpgradeAdd2Outlook.Add_Click( { Start-Process PowerShell .\Auto_Upgrade_Add2Outlook })
$O365ExchangePermissions.Add_Click( { Start-Process Powershell .\PermissionsOnPremOrO365Combined.ps1 })
$A2OPermissions.Add_Click( { Start-Process Powershell .\Add2Outlook_Set_Granular_permissions.ps1 })
$AutoPermissions.Add_Click( { Start-Process PowerShell .\Permissions_Task_Creation.ps1 })
$DownloadAdd2Exchange.Add_Click( { Start-Process http://support.DidItBetter.com/Secure/Login.aspx?returnurl=/downloads.aspx })
$DownloadToolKit.Add_Click( { Start-Process ftp://ftp.diditbetter.com/Add2Outlook%20Toolkit/Upgrades/Add2Outlook%20ToolKit%20Full%20Installation.exe })
$DownloadSQL.Add_Click( { Start-Process https://aka.ms/ssmsfullsetup })
$A2EDiags.Add_Click( { Start-Process PowerShell .\Get_Diags.ps1 })
$FTPDownloads.Add_Click( { Start-Process ftp.DidItBetter.com })
$GetHelp.Add_Click( { Start-Process http://support.DidItBetter.com/support-request.aspx })
$SearchDidItBetter.Add_Click( { Start-Process http://support.DidItBetter.com/ })
$QuickStartGuide.Add_Click( { Start-Process http://guides.diditbetter.com/Add2Exchange_Guide.pdf })
$AutoLogon.Add_Click( { Start-Process Powershell .\AutoLogon.exe })
$DirSync.Add_Click( { Start-Process Powershell .\Dir_Sync.ps1 })
$DisableUAC.Add_Click( { Start-Process Powershell .\Disable_UAC.ps1 })
$GroupPolicyResults.Add_Click( { Start-Process Powershell .\GP_Results.ps1 })
$CheckPowerShell.Add_Click( { Start-Process Powershell .\Legacy_PowerShell.ps1 })
$DisableOSC.Add_Click( {
        Copy-Item ".\OSC_Disable.bat" -Destination "$Home\Desktop\" -ErrorAction SilentlyContinue
        Start-Process Powershell .\OSC_Disable.bat })
$AddRegistryFavorites.Add_Click( { Start-Process Powershell .\Registry_Favorites.ps1 })
$Reset_A2E_Passwords.Add_Click( { Start-Process Powershell .\Reset_A2E_Password.ps1 })
$CreateSupporttext.Add_Click( { Start-Process PowerShell .\A2E_Setup_Details.ps1 })
$GALSync.Add_Click( { Start-Process http://guides.diditbetter.com/GAL_Sync_Scenario.pdf })
$PrivatetoPrivate.Add_Click( { Start-Process http://guides.diditbetter.com/Private_to_Private_Sync_Scenarios.pdf })
$PublictoPublic.Add_Click( { Start-Process http://guides.diditbetter.com/Public_to_Public_Sync_Scenarios.pdf })
$PrivatetoPublic.Add_Click( { Start-Process http://guides.diditbetter.com/Private_to_Public_Sync_Scenarios.pdf })
$PublictoPrivate.Add_Click( { Start-Process http://guides.diditbetter.com/Public_to_Private_Sync_Scenarios.pdf })
$TemplateCreation.Add_Click( { Start-Process http://guides.diditbetter.com/Template_Creation_RGM_Sync_Scenarios.pdf })
$MigrateA2E.Add_Click( { Start-Process http://guides.diditbetter.com/Migrating_A2E_Sync_Scenarios.pdf })
$ExhangeMigration.Add_Click( { Start-Process http://guides.diditbetter.com/Migrating_Environments_A2E_Sync_Scenarios.pdf })
$AD_Photos.Add_Click( { Start-Process Powershell .\Export_ADPhoto.ps1 })
$MSExchangeDelegate.Add_Click( { Start-Process Powershell .\MSExchangeDelegation.ps1 })
$ExchangeShell.Add_Click( { Start-Process Powershell .\Shell.ps1 })
$CommandsList.Add_Click( { Invoke-Item .\A2E_Permissions_Commands.rtf })
$A2E_Migration_Wizard.Add_Click( { Start-Process Powershell .\A2E_Auto_Migration.ps1 })
$ModernAuth.Add_Click( { Start-Process Powershell .\Disable_Modern_Authentication.ps1 })
$DIBMMC.Add_Click({Start-Process Powershell .\A2E_MMC.ps1})
$RegistryEditor.Add_Click({Start-Process Regedit})
$DIBDirectory.Add_Click({Start-Process Powershell .\A2E_Directory.ps1})
$TaskScheduler.Add_Click({Start-Process taskschd.msc})
$WindowsDefender.Add_Click({Start-Process Powershell .\Windows_Defender_Exclusions.ps1})
$Outlook_Tools_Menu.Add_Click({Start-Process Powershell .\Outlook_Tools_Menu.ps1})
$A2ESQLBackUp.Add_Click({Start-Process Powershell .\A2E_SQL_Backup.ps1})
$SQLUpgrade.Add_Click({Start-Process Powershell .\Upgrade_SQL_2008.ps1})



[void]$DidItBetterSupportMenu.ShowDialog()