<#
.NAME
    A2EMenu
.DESCRIPTION
    Add2Exchange Menu
#>


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$DidItBetterSupportMenu          = New-Object system.Windows.Forms.Form
$DidItBetterSupportMenu.ClientSize  = '547,596'
$DidItBetterSupportMenu.text     = "DidItBetter Software Support Menu"
$DidItBetterSupportMenu.BackColor  = "#ffffff"
$DidItBetterSupportMenu.TopMost  = $false

$Upgrades                        = New-Object system.Windows.Forms.Label
$Upgrades.text                   = "Upgrades"
$Upgrades.AutoSize               = $true
$Upgrades.width                  = 150
$Upgrades.height                 = 10
$Upgrades.location               = New-Object System.Drawing.Point(11,10)
$Upgrades.Font                   = 'Microsoft Sans Serif,12,style=Bold,Underline'

$Add2ExchangeUpgrade             = New-Object system.Windows.Forms.Label
$Add2ExchangeUpgrade.text        = "Upgrade Add2Exchange Enterprise"
$Add2ExchangeUpgrade.AutoSize    = $true
$Add2ExchangeUpgrade.width       = 150
$Add2ExchangeUpgrade.height      = 10
$Add2ExchangeUpgrade.location    = New-Object System.Drawing.Point(11,35)
$Add2ExchangeUpgrade.Font        = 'Microsoft Sans Serif,9'

$UpgradeRMM                      = New-Object system.Windows.Forms.Label
$UpgradeRMM.text                 = "Upgrade Recovery and Migration Manager"
$UpgradeRMM.AutoSize             = $true
$UpgradeRMM.width                = 150
$UpgradeRMM.height               = 10
$UpgradeRMM.location             = New-Object System.Drawing.Point(11,55)
$UpgradeRMM.Font                 = 'Microsoft Sans Serif,9'

$AddingPermissions               = New-Object system.Windows.Forms.Label
$AddingPermissions.text          = "Adding Permissions"
$AddingPermissions.AutoSize      = $true
$AddingPermissions.width         = 25
$AddingPermissions.height        = 10
$AddingPermissions.location      = New-Object System.Drawing.Point(11,146)
$AddingPermissions.Font          = 'Microsoft Sans Serif,12,style=Bold,Underline'

$O365ExchangePermissions         = New-Object system.Windows.Forms.Label
$O365ExchangePermissions.text    = "Office 365 and On-Premise Exchange Permissions"
$O365ExchangePermissions.AutoSize  = $true
$O365ExchangePermissions.width   = 150
$O365ExchangePermissions.height  = 10
$O365ExchangePermissions.location  = New-Object System.Drawing.Point(11,171)
$O365ExchangePermissions.Font    = 'Microsoft Sans Serif,9'

$A2OPermissions                  = New-Object system.Windows.Forms.Label
$A2OPermissions.text             = "Add2Outlook Granular Permissions"
$A2OPermissions.AutoSize         = $true
$A2OPermissions.width            = 150
$A2OPermissions.height           = 10
$A2OPermissions.location         = New-Object System.Drawing.Point(11,191)
$A2OPermissions.Font             = 'Microsoft Sans Serif,9'

$AutoPermissions                 = New-Object system.Windows.Forms.Label
$AutoPermissions.text            = "Automate Permissions on a Schedule"
$AutoPermissions.AutoSize        = $true
$AutoPermissions.width           = 150
$AutoPermissions.height          = 10
$AutoPermissions.location        = New-Object System.Drawing.Point(11,211)
$AutoPermissions.Font            = 'Microsoft Sans Serif,9'

$Downloads                       = New-Object system.Windows.Forms.Label
$Downloads.text                  = "Downloads"
$Downloads.AutoSize              = $true
$Downloads.width                 = 150
$Downloads.height                = 10
$Downloads.location              = New-Object System.Drawing.Point(295,10)
$Downloads.Font                  = 'Microsoft Sans Serif,12,style=Bold,Underline'

$DownloadAdd2Exchange            = New-Object system.Windows.Forms.Label
$DownloadAdd2Exchange.text       = "Download Add2Exchange"
$DownloadAdd2Exchange.AutoSize   = $true
$DownloadAdd2Exchange.width      = 150
$DownloadAdd2Exchange.height     = 10
$DownloadAdd2Exchange.location   = New-Object System.Drawing.Point(295,35)
$DownloadAdd2Exchange.Font       = 'Microsoft Sans Serif,9'

$DownloadToolKit                 = New-Object system.Windows.Forms.Label
$DownloadToolKit.text            = "Download ToolKit"
$DownloadToolKit.AutoSize        = $true
$DownloadToolKit.width           = 150
$DownloadToolKit.height          = 10
$DownloadToolKit.location        = New-Object System.Drawing.Point(295,55)
$DownloadToolKit.Font            = 'Microsoft Sans Serif,9'

$DownloadSQL                     = New-Object system.Windows.Forms.Label
$DownloadSQL.text                = "Download SQL Studio"
$DownloadSQL.AutoSize            = $true
$DownloadSQL.width               = 150
$DownloadSQL.height              = 10
$DownloadSQL.location            = New-Object System.Drawing.Point(295,75)
$DownloadSQL.Font                = 'Microsoft Sans Serif,9'

$FTPDownloads                    = New-Object system.Windows.Forms.Label
$FTPDownloads.text               = "FTP Downloads"
$FTPDownloads.AutoSize           = $true
$FTPDownloads.width              = 150
$FTPDownloads.height             = 10
$FTPDownloads.location           = New-Object System.Drawing.Point(295,115)
$FTPDownloads.Font               = 'Microsoft Sans Serif,9'

$GetSupport                      = New-Object system.Windows.Forms.Label
$GetSupport.text                 = "Get Support"
$GetSupport.AutoSize             = $true
$GetSupport.width                = 150
$GetSupport.height               = 10
$GetSupport.location             = New-Object System.Drawing.Point(295,146)
$GetSupport.Font                 = 'Microsoft Sans Serif,12,style=Bold,Underline'

$GetHelp                         = New-Object system.Windows.Forms.Label
$GetHelp.text                    = "Need Help? Open a Ticket!"
$GetHelp.AutoSize                = $true
$GetHelp.width                   = 150
$GetHelp.height                  = 10
$GetHelp.location                = New-Object System.Drawing.Point(295,171)
$GetHelp.Font                    = 'Microsoft Sans Serif,9'

$SearchDidItBetter               = New-Object system.Windows.Forms.Label
$SearchDidItBetter.text          = "Search DidItBetter"
$SearchDidItBetter.AutoSize      = $true
$SearchDidItBetter.width         = 150
$SearchDidItBetter.height        = 10
$SearchDidItBetter.location      = New-Object System.Drawing.Point(295,191)
$SearchDidItBetter.Font          = 'Microsoft Sans Serif,9'

$QuickStartGuide                 = New-Object system.Windows.Forms.Label
$QuickStartGuide.text            = "Quick Start Guide"
$QuickStartGuide.AutoSize        = $true
$QuickStartGuide.width           = 150
$QuickStartGuide.height          = 10
$QuickStartGuide.location        = New-Object System.Drawing.Point(295,211)
$QuickStartGuide.Font            = 'Microsoft Sans Serif,9'

$Tools                           = New-Object system.Windows.Forms.Label
$Tools.text                      = "Tools"
$Tools.AutoSize                  = $true
$Tools.width                     = 150
$Tools.height                    = 10
$Tools.location                  = New-Object System.Drawing.Point(11,251)
$Tools.Font                      = 'Microsoft Sans Serif,12,style=Bold,Underline'

$AutoLogon                       = New-Object system.Windows.Forms.Label
$AutoLogon.text                  = "Auto Logon"
$AutoLogon.AutoSize              = $true
$AutoLogon.width                 = 150
$AutoLogon.height                = 10
$AutoLogon.location              = New-Object System.Drawing.Point(11,276)
$AutoLogon.Font                  = 'Microsoft Sans Serif,9'

$DirSync                         = New-Object system.Windows.Forms.Label
$DirSync.text                    = "Dir Sync"
$DirSync.AutoSize                = $true
$DirSync.width                   = 150
$DirSync.height                  = 10
$DirSync.location                = New-Object System.Drawing.Point(11,296)
$DirSync.Font                    = 'Microsoft Sans Serif,9'

$DisableUAC                      = New-Object system.Windows.Forms.Label
$DisableUAC.text                 = "Disable User Account Control"
$DisableUAC.AutoSize             = $true
$DisableUAC.width                = 150
$DisableUAC.height               = 10
$DisableUAC.location             = New-Object System.Drawing.Point(11,316)
$DisableUAC.Font                 = 'Microsoft Sans Serif,9'

$GroupPolicyResults              = New-Object system.Windows.Forms.Label
$GroupPolicyResults.text         = "Group Policy Results"
$GroupPolicyResults.AutoSize     = $true
$GroupPolicyResults.width        = 150
$GroupPolicyResults.height       = 10
$GroupPolicyResults.location     = New-Object System.Drawing.Point(11,336)
$GroupPolicyResults.Font         = 'Microsoft Sans Serif,9'

$CheckPowerShell                 = New-Object system.Windows.Forms.Label
$CheckPowerShell.text            = "Check and Upgrade PowerShell"
$CheckPowerShell.AutoSize        = $true
$CheckPowerShell.width           = 150
$CheckPowerShell.height          = 10
$CheckPowerShell.location        = New-Object System.Drawing.Point(11,356)
$CheckPowerShell.Font            = 'Microsoft Sans Serif,9'

$RemoveOutlookAddons             = New-Object system.Windows.Forms.Label
$RemoveOutlookAddons.text        = "Remove Outlook Add-ins"
$RemoveOutlookAddons.AutoSize    = $true
$RemoveOutlookAddons.width       = 150
$RemoveOutlookAddons.height      = 10
$RemoveOutlookAddons.location    = New-Object System.Drawing.Point(11,376)
$RemoveOutlookAddons.Font        = 'Microsoft Sans Serif,9'

$IncludeRegistryFavorites        = New-Object system.Windows.Forms.Label
$IncludeRegistryFavorites.text   = "Include Registry Favorites"
$IncludeRegistryFavorites.AutoSize  = $true
$IncludeRegistryFavorites.width  = 150
$IncludeRegistryFavorites.height  = 10
$IncludeRegistryFavorites.location  = New-Object System.Drawing.Point(11,396)
$IncludeRegistryFavorites.Font   = 'Microsoft Sans Serif,9'

$ExportLicenseandProfile1        = New-Object system.Windows.Forms.Label
$ExportLicenseandProfile1.text   = "Export Add2Exchange License and Profile 1"
$ExportLicenseandProfile1.AutoSize  = $true
$ExportLicenseandProfile1.width  = 150
$ExportLicenseandProfile1.height  = 10
$ExportLicenseandProfile1.location  = New-Object System.Drawing.Point(11,416)
$ExportLicenseandProfile1.Font   = 'Microsoft Sans Serif,9'

$SyncScenarios                   = New-Object system.Windows.Forms.Label
$SyncScenarios.text              = "Add2Exchange Sync Scenarios"
$SyncScenarios.AutoSize          = $true
$SyncScenarios.width             = 150
$SyncScenarios.height            = 10
$SyncScenarios.location          = New-Object System.Drawing.Point(295,251)
$SyncScenarios.Font              = 'Microsoft Sans Serif,12,style=Bold,Underline'

$GALSync                         = New-Object system.Windows.Forms.Label
$GALSync.text                    = "GAL Synchronization"
$GALSync.AutoSize                = $true
$GALSync.width                   = 150
$GALSync.height                  = 10
$GALSync.location                = New-Object System.Drawing.Point(295,276)
$GALSync.Font                    = 'Microsoft Sans Serif,9'

$PrivatetoPrivate                = New-Object system.Windows.Forms.Label
$PrivatetoPrivate.text           = "Private to Private Relationship"
$PrivatetoPrivate.AutoSize       = $true
$PrivatetoPrivate.width          = 150
$PrivatetoPrivate.height         = 10
$PrivatetoPrivate.location       = New-Object System.Drawing.Point(295,296)
$PrivatetoPrivate.Font           = 'Microsoft Sans Serif,9'

$PublictoPublic                  = New-Object system.Windows.Forms.Label
$PublictoPublic.text             = "Public to Public Relationship"
$PublictoPublic.AutoSize         = $true
$PublictoPublic.width            = 150
$PublictoPublic.height           = 10
$PublictoPublic.location         = New-Object System.Drawing.Point(295,316)
$PublictoPublic.Font             = 'Microsoft Sans Serif,9'

$PrivatetoPublic                 = New-Object system.Windows.Forms.Label
$PrivatetoPublic.text            = "Private to Public Relationship"
$PrivatetoPublic.AutoSize        = $true
$PrivatetoPublic.width           = 150
$PrivatetoPublic.height          = 10
$PrivatetoPublic.location        = New-Object System.Drawing.Point(295,336)
$PrivatetoPublic.Font            = 'Microsoft Sans Serif,9'

$PublictoPrivate                 = New-Object system.Windows.Forms.Label
$PublictoPrivate.text            = "Public to Private Relationship"
$PublictoPrivate.AutoSize        = $true
$PublictoPrivate.width           = 150
$PublictoPrivate.height          = 10
$PublictoPrivate.location        = New-Object System.Drawing.Point(295,356)
$PublictoPrivate.Font            = 'Microsoft Sans Serif,9'

$TemplateCreation                = New-Object system.Windows.Forms.Label
$TemplateCreation.text           = "Template Creation"
$TemplateCreation.AutoSize       = $true
$TemplateCreation.width          = 150
$TemplateCreation.height         = 10
$TemplateCreation.location       = New-Object System.Drawing.Point(295,376)
$TemplateCreation.Font           = 'Microsoft Sans Serif,9'

$MIgrateA2E                      = New-Object system.Windows.Forms.Label
$MIgrateA2E.text                 = "Migrate Add2Exchange to a New Box"
$MIgrateA2E.AutoSize             = $true
$MIgrateA2E.width                = 150
$MIgrateA2E.height               = 10
$MIgrateA2E.location             = New-Object System.Drawing.Point(295,396)
$MIgrateA2E.Font                 = 'Microsoft Sans Serif,9'

$ExhangeMigration                = New-Object system.Windows.Forms.Label
$ExhangeMigration.text           = "Exchange Migration with Add2Exchange"
$ExhangeMigration.AutoSize       = $true
$ExhangeMigration.width          = 150
$ExhangeMigration.height         = 10
$ExhangeMigration.location       = New-Object System.Drawing.Point(295,416)
$ExhangeMigration.Font           = 'Microsoft Sans Serif,9'

$DidItBetterLogo                 = New-Object system.Windows.Forms.PictureBox
$DidItBetterLogo.width           = 201
$DidItBetterLogo.height          = 86
$DidItBetterLogo.location        = New-Object System.Drawing.Point(14,495)
$DidItBetterLogo.imageLocation   = "./Diditbetter_logo.png"
$DidItBetterLogo.SizeMode        = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$UpgradeToolkit                  = New-Object system.Windows.Forms.Label
$UpgradeToolkit.text             = "Upgrade Add2Outlook ToolKit"
$UpgradeToolkit.AutoSize         = $true
$UpgradeToolkit.width            = 150
$UpgradeToolkit.height           = 10
$UpgradeToolkit.location         = New-Object System.Drawing.Point(11,75)
$UpgradeToolkit.Font             = 'Microsoft Sans Serif,9'

$UpgradeAdd2Outlook              = New-Object system.Windows.Forms.Label
$UpgradeAdd2Outlook.text         = "Upgrade Add2Outlook"
$UpgradeAdd2Outlook.AutoSize     = $true
$UpgradeAdd2Outlook.width        = 25
$UpgradeAdd2Outlook.height       = 10
$UpgradeAdd2Outlook.location     = New-Object System.Drawing.Point(11,95)
$UpgradeAdd2Outlook.Font         = 'Microsoft Sans Serif,9'

$A2EDiags                        = New-Object system.Windows.Forms.Label
$A2EDiags.text                   = "Get A2E Diags"
$A2EDiags.AutoSize               = $true
$A2EDiags.width                  = 150
$A2EDiags.height                 = 10
$A2EDiags.location               = New-Object System.Drawing.Point(295,95)
$A2EDiags.Font                   = 'Microsoft Sans Serif,9'

$CreateSupporttext               = New-Object system.Windows.Forms.Label
$CreateSupporttext.text          = "Create Support Setup Details"
$CreateSupporttext.AutoSize      = $true
$CreateSupporttext.width         = 150
$CreateSupporttext.height        = 10
$CreateSupporttext.location      = New-Object System.Drawing.Point(11,436)
$CreateSupporttext.Font          = 'Microsoft Sans Serif,9'

$Revision                        = New-Object system.Windows.Forms.Label
$Revision.text                   = "Rev. 1.31819"
$Revision.AutoSize               = $true
$Revision.width                  = 25
$Revision.height                 = 10
$Revision.location               = New-Object System.Drawing.Point(445,572)
$Revision.Font                   = 'Microsoft Sans Serif,8,style=Italic'

$ToolTip1                        = New-Object system.Windows.Forms.ToolTip
$ToolTip1.ToolTipTitle           = "Help"
$ToolTip1.isBalloon              = $true

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
$ToolTip1.SetToolTip($RemoveOutlookAddons,'PowerShell Script to Remove Outlook Add-Ins like Social Connector')
$ToolTip1.SetToolTip($IncludeRegistryFavorites,'PowerShell Script to Add Add2Exchange Favorites in the Registry')
$ToolTip1.SetToolTip($ExportLicenseandProfile1,'PowerShell Script to Export and Save your License and User Information for Add2Exchange')
$ToolTip1.SetToolTip($GALSync,'Click for more information on How To Setup a Global Address Sync ')
$ToolTip1.SetToolTip($PrivatetoPrivate,'Click for more information on How To Setup a Private to Private Relationship')
$ToolTip1.SetToolTip($PublictoPublic,'Click for more information on How To Setup a Public to Public Relationship')
$ToolTip1.SetToolTip($PrivatetoPublic,'Click for more information on How To Setup a Private to Public')
$ToolTip1.SetToolTip($PublictoPrivate,'Click for more information on How To Setup a Public to Private')
$ToolTip1.SetToolTip($TemplateCreation,'Click for more information on How To Setup Templates for Synchronization with Distribution Groups')
$ToolTip1.SetToolTip($MIgrateA2E,'Click for more information on How To Migrate Add2Exchange onto a new Appliance')
$ToolTip1.SetToolTip($ExhangeMigration,'Click for more information on How To Setup Add2Exchange before or After an Exchange or Office 365 Migration')
$ToolTip1.SetToolTip($UpgradeToolkit,'This will Upgrade your Current Version of Add2Outlook Toolkit to the Latest Version')
$ToolTip1.SetToolTip($CreateSupporttext,'PowerShell Script to Create a Support text detailing this Installation')
$ToolTip1.SetToolTip($UpgradeAdd2Outlook,'This will Upgrade your Current Version of Add2Outlook to the Latest Version')

$DidItBetterSupportMenu.controls.AddRange(@($Upgrades,$Add2ExchangeUpgrade,$UpgradeRMM,$UpgradeToolkit,$AddingPermissions,$O365ExchangePermissions,$A2OPermissions,$AutoPermissions,$Downloads,$DownloadAdd2Exchange,$DownloadToolKit,$DownloadSQL,$FTPDownloads,
$GetSupport,$GetHelp,$SearchDidItBetter,$QuickStartGuide,$Tools,$AutoLogon,$DirSync,$DisableUAC,$GroupPolicyResults,$CheckPowerShell,$RemoveOutlookAddons,$IncludeRegistryFavorites,$ExportLicenseandProfile1,
$SyncScenarios,$GALSync,$PrivatetoPrivate,$PublictoPublic,$PrivatetoPublic,$PublictoPrivate,$TemplateCreation,$MigrateA2E,$ExhangeMigration,$DidItBetterLogo,$A2EDiags,$CreateSupporttext,$UpgradeAdd2Outlook,$Revision))

$Add2ExchangeUpgrade.Add_Click({Start-Process Powershell .\Auto_Upgrade_Add2Exchange.ps1})
$UpgradeRMM.Add_Click({Start-Process Powershell .\Auto_Upgrade_RMM.ps1})
$UpgradeToolkit.Add_Click({Start-Process Powershell .\Auto_Upgrade_ToolKit.ps1})
$UpgradeAdd2Outlook.Add_Click({Start-Process PowerShell .\Auto_Upgrade_Add2Outlook})
$O365ExchangePermissions.Add_Click({Start-Process Powershell .\PermissionsOnPremOrO365Combined.ps1})
$A2OPermissions.Add_Click({Start-Process Powershell .\Add2Outlook_Set_Granular_permissions.ps1})
$AutoPermissions.Add_Click({Start-Process PowerShell .\Permissions_Task_Creation.ps1})
$DownloadAdd2Exchange.Add_Click({Start-Process http://support.DidItBetter.com/Secure/Login.aspx?returnurl=/downloads.aspx})
$DownloadToolKit.Add_Click({Start-Process ftp://ftp.diditbetter.com/Add2Outlook%20Toolkit/Upgrades/Add2Outlook%20ToolKit%20Full%20Installation.exe})
$DownloadSQL.Add_Click({Start-Process ftp://ftp.DidItBetter.com/SQL/SQL2012Management/2012SQLManagementStudio_x64_ENU.exe})
$A2EDiags.Add_Click({Start-Process PowerShell .\Get_Diags.ps1})
$FTPDownloads.Add_Click({Start-Process ftp.DidItBetter.com})
$GetHelp.Add_Click({Start-Process http://support.DidItBetter.com/support-request.aspx})
$SearchDidItBetter.Add_Click({Start-Process http://support.DidItBetter.com/})
$QuickStartGuide.Add_Click({Start-Process http://guides.diditbetter.com/Add2Exchange_Guide.pdf})
$AutoLogon.Add_Click({Start-Process Powershell .\AutoLogon.exe})
$DirSync.Add_Click({Start-Process Powershell .\Dir_Sync.ps1})
$DisableUAC.Add_Click({Start-Process Powershell .\Disable_UAC.ps1})
$GroupPolicyResults.Add_Click({Start-Process Powershell .\GP_Results.ps1})
$CheckPowerShell.Add_Click({Start-Process Powershell .\Legacy_PowerShell.ps1})
$RemoveOutlookAddons.Add_Click({Start-Process Powershell .\Remove_Outlook_Add_ins.ps1})
$IncludeRegistryFavorites.Add_Click({Start-Process Powershell.\Registry_Favorites.ps1})
$ExportLicenseandProfile1.Add_Click({Start-Process Powershell.\Export_License_and_Profile1.ps1})
$CreateSupporttext.Add_Click({Start-Process PowerShell .\A2E_Setup_Details.ps1})
$GALSync.Add_Click({Start-Process http://guides.diditbetter.com/GAL_Sync_Scenario.pdf})
$PrivatetoPrivate.Add_Click({Start-Process http://guides.diditbetter.com/Private_to_Private_Sync_Scenarios.pdf})
$PublictoPublic.Add_Click({Start-Process http://guides.diditbetter.com/Public_to_Public_Sync_Scenarios.pdf})
$PrivatetoPublic.Add_Click({Start-Process http://guides.diditbetter.com/Private_to_Public_Sync_Scenarios.pdf})
$PublictoPrivate.Add_Click({Start-Process http://guides.diditbetter.com/Public_to_Private_Sync_Scenarios.pdf})
$TemplateCreation.Add_Click({Start-Process http://guides.diditbetter.com/Template_Creation_RGM_Sync_Scenarios.pdf})
$MigrateA2E.Add_Click({Start-Process http://guides.diditbetter.com/Migrating_A2E_Sync_Scenarios.pdf})
$ExhangeMigration.Add_Click({Start-Process http://guides.diditbetter.com/Migrating_Environments_A2E_Sync_Scenarios.pdf})



#Write your logic code here

[void]$DidItBetterSupportMenu.ShowDialog()