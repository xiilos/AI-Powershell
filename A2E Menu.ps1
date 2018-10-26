if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Unrestricted


<#
.NAME
    Add2Exchange ToolBox
.SYNOPSIS
    Add2Exchange Tool Box
.DESCRIPTION
    Tools to install and Administer Add2Exchange
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$A2ETools                        = New-Object system.Windows.Forms.Form
$A2ETools.ClientSize             = '710,598'
$A2ETools.text                   = "Add2Exchange Toolbox"
$A2ETools.BackColor              = "#ffffff"
$A2ETools.TopMost                = $false

$ExchangePermissions             = New-Object system.Windows.Forms.Panel
$ExchangePermissions.height      = 205
$ExchangePermissions.width       = 310
$ExchangePermissions.location    = New-Object System.Drawing.Point(19,50)

$Add2Exchange                    = New-Object system.Windows.Forms.Panel
$Add2Exchange.height             = 205
$Add2Exchange.width              = 310
$Add2Exchange.location           = New-Object System.Drawing.Point(380,50)

$WindowsTasks                    = New-Object system.Windows.Forms.Panel
$WindowsTasks.height             = 280
$WindowsTasks.width              = 310
$WindowsTasks.location           = New-Object System.Drawing.Point(19,299)

$O365onprem                      = New-Object system.Windows.Forms.Button
$O365onprem.text                 = "Office 365 & On-Premise Permissions"
$O365onprem.width                = 250
$O365onprem.height               = 30
$O365onprem.location             = New-Object System.Drawing.Point(10,50)
$O365onprem.Font                 = 'Microsoft Sans Serif,10'

$UAC                             = New-Object system.Windows.Forms.Button
$UAC.text                        = "Disable UAC"
$UAC.width                       = 175
$UAC.height                      = 25
$UAC.location                    = New-Object System.Drawing.Point(15,20)
$UAC.Font                        = 'Microsoft Sans Serif,10'

$Favorites                       = New-Object system.Windows.Forms.Button
$Favorites.text                  = "Add Add2Exchange Favorites"
$Favorites.width                 = 220
$Favorites.height                = 30
$Favorites.location              = New-Object System.Drawing.Point(10,10)
$Favorites.Font                  = 'Microsoft Sans Serif,10'

$Add2Outlook                     = New-Object system.Windows.Forms.Button
$Add2Outlook.text                = "A2O Granular Permissions"
$Add2Outlook.width               = 220
$Add2Outlook.height              = 30
$Add2Outlook.location            = New-Object System.Drawing.Point(10,50)
$Add2Outlook.Font                = 'Microsoft Sans Serif,10'

$AutoLogin                       = New-Object system.Windows.Forms.Button
$AutoLogin.text                  = "Setup Auto Login"
$AutoLogin.width                 = 175
$AutoLogin.height                = 25
$AutoLogin.location              = New-Object System.Drawing.Point(15,60)
$AutoLogin.Font                  = 'Microsoft Sans Serif,10'

$A2EExport                       = New-Object system.Windows.Forms.Button
$A2EExport.text                  = "Export License and Profile 1 Info"
$A2EExport.width                 = 220
$A2EExport.height                = 30
$A2EExport.location              = New-Object System.Drawing.Point(10,90)
$A2EExport.Font                  = 'Microsoft Sans Serif,10'

$GPOResult                       = New-Object system.Windows.Forms.Button
$GPOResult.text                  = "Get Group Policy Results"
$GPOResult.width                 = 175
$GPOResult.height                = 25
$GPOResult.location              = New-Object System.Drawing.Point(15,100)
$GPOResult.Font                  = 'Microsoft Sans Serif,10'

$ShellUpdate                     = New-Object system.Windows.Forms.Button
$ShellUpdate.text                = "Get MS Online Azure AD & Update"
$ShellUpdate.width               = 250
$ShellUpdate.height              = 30
$ShellUpdate.location            = New-Object System.Drawing.Point(10,10)
$ShellUpdate.Font                = 'Microsoft Sans Serif,10'

$OutlookAddinDisable             = New-Object system.Windows.Forms.Button
$OutlookAddinDisable.text        = "Disable Outlook Add-ins"
$OutlookAddinDisable.width       = 175
$OutlookAddinDisable.height      = 25
$OutlookAddinDisable.location    = New-Object System.Drawing.Point(15,140)
$OutlookAddinDisable.Font        = 'Microsoft Sans Serif,10'

$OutlookAddinsEnable             = New-Object system.Windows.Forms.Button
$OutlookAddinsEnable.text        = "Enable Outlook Add-ins"
$OutlookAddinsEnable.width       = 175
$OutlookAddinsEnable.height      = 25
$OutlookAddinsEnable.location    = New-Object System.Drawing.Point(16,180)
$OutlookAddinsEnable.Font        = 'Microsoft Sans Serif,10'

$Windows10Virgin                 = New-Object system.Windows.Forms.Button
$Windows10Virgin.text            = "Virginize Windows 10"
$Windows10Virgin.width           = 175
$Windows10Virgin.height          = 25
$Windows10Virgin.location        = New-Object System.Drawing.Point(15,220)
$Windows10Virgin.Font            = 'Microsoft Sans Serif,10'

$ExchangeSideLabel               = New-Object system.Windows.Forms.Label
$ExchangeSideLabel.text          = "Exchange Services"
$ExchangeSideLabel.AutoSize      = $true
$ExchangeSideLabel.width         = 25
$ExchangeSideLabel.height        = 10
$ExchangeSideLabel.location      = New-Object System.Drawing.Point(18,25)
$ExchangeSideLabel.Font          = 'Microsoft Sans Serif,10,style=Bold'

$A2ELable                        = New-Object system.Windows.Forms.Label
$A2ELable.text                   = "Add2Exchange Services"
$A2ELable.AutoSize               = $true
$A2ELable.width                  = 25
$A2ELable.height                 = 10
$A2ELable.location               = New-Object System.Drawing.Point(380,25)
$A2ELable.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$WindowsServices                 = New-Object system.Windows.Forms.Label
$WindowsServices.text            = "Windows Services"
$WindowsServices.AutoSize        = $true
$WindowsServices.width           = 25
$WindowsServices.height          = 10
$WindowsServices.location        = New-Object System.Drawing.Point(19,275)
$WindowsServices.Font            = 'Microsoft Sans Serif,10,style=Bold'

$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 280
$Panel1.width                    = 310
$Panel1.location                 = New-Object System.Drawing.Point(380,298)

$A2ESQL                          = New-Object system.Windows.Forms.Label
$A2ESQL.text                     = "Add2Exchange SQL"
$A2ESQL.AutoSize                 = $true
$A2ESQL.width                    = 25
$A2ESQL.height                   = 10
$A2ESQL.location                 = New-Object System.Drawing.Point(380,275)
$A2ESQL.Font                     = 'Microsoft Sans Serif,10,style=Bold'

$SQLCommands                     = New-Object system.Windows.Forms.Button
$SQLCommands.text                = "SQL Coming soon!"
$SQLCommands.width               = 230
$SQLCommands.height              = 163
$SQLCommands.location            = New-Object System.Drawing.Point(42,50)
$SQLCommands.Font                = 'Microsoft Sans Serif,10'

$A2Eupdates                      = New-Object system.Windows.Forms.Button
$A2Eupdates.text                 = "Update Add2Exchange"
$A2Eupdates.width                = 220
$A2Eupdates.height               = 30
$A2Eupdates.location             = New-Object System.Drawing.Point(10,130)
$A2Eupdates.Font                 = 'Microsoft Sans Serif,10'

$A2ETools.controls.AddRange(@($ExchangePermissions,$Add2Exchange,$WindowsTasks,$ExchangeSideLabel,$A2ELable,$WindowsServices,$Panel1,$A2ESQL))
$ExchangePermissions.controls.AddRange(@($O365onprem,$ShellUpdate))
$WindowsTasks.controls.AddRange(@($UAC,$AutoLogin,$GPOResult,$OutlookAddinDisable,$OutlookAddinsEnable,$Windows10Virgin))
$Add2Exchange.controls.AddRange(@($Favorites,$Add2Outlook,$A2EExport,$A2Eupdates))
$Panel1.controls.AddRange(@($SQLCommands))

#region gui events {
$O365onprem.Add_Click({  })
$ShellUpdate.Add_Click({  })
$Favorites.Add_Click({  })
$Add2Outlook.Add_Click({  })
$A2EExport.Add_Click({  })
$UAC.Add_Click({  })
$AutoLogin.Add_Click({  })
$GPOResult.Add_Click({  })
$OutlookAddinDisable.Add_Click({  })
$OutlookAddinsEnable.Add_Click({  })
$Windows10Virgin.Add_Click({  })
#endregion events }

#endregion GUI }


# Logic Code Here

[void]$A2ETools.ShowDialog()




Write-Output "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting