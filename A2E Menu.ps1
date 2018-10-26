if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}
# Execution Policy
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
$A2ETools.ClientSize             = '596,573'
$A2ETools.text                   = "Add2Exchange Toolbox"
$A2ETools.BackColor              = "#ffffff"
$A2ETools.TopMost                = $false

$ExchangePermissions             = New-Object system.Windows.Forms.Panel
$ExchangePermissions.height      = 205
$ExchangePermissions.width       = 270
$ExchangePermissions.location    = New-Object System.Drawing.Point(15,61)

$Add2Exchange                    = New-Object system.Windows.Forms.Panel
$Add2Exchange.height             = 205
$Add2Exchange.width              = 270
$Add2Exchange.location           = New-Object System.Drawing.Point(307,61)

$WindowsTasks                    = New-Object system.Windows.Forms.Panel
$WindowsTasks.height             = 280
$WindowsTasks.width              = 270
$WindowsTasks.location           = New-Object System.Drawing.Point(14,279)

$O365onprem                      = New-Object system.Windows.Forms.Button
$O365onprem.text                 = "Office 365 or On-Premise Permissions"
$O365onprem.width                = 250
$O365onprem.height               = 30
$O365onprem.location             = New-Object System.Drawing.Point(10,70)
$O365onprem.Font                 = 'Microsoft Sans Serif,10'

$UAC                             = New-Object system.Windows.Forms.Button
$UAC.text                        = "Disable UAC"
$UAC.width                       = 175
$UAC.height                      = 25
$UAC.location                    = New-Object System.Drawing.Point(7,30)
$UAC.Font                        = 'Microsoft Sans Serif,10'

$Favorites                       = New-Object system.Windows.Forms.Button
$Favorites.text                  = "Add Add2Exchange Favorites"
$Favorites.width                 = 220
$Favorites.height                = 30
$Favorites.location              = New-Object System.Drawing.Point(10,30)
$Favorites.Font                  = 'Microsoft Sans Serif,10'

$Add2Outlook                     = New-Object system.Windows.Forms.Button
$Add2Outlook.text                = "A2O Granular Permissions"
$Add2Outlook.width               = 220
$Add2Outlook.height              = 30
$Add2Outlook.location            = New-Object System.Drawing.Point(10,70)
$Add2Outlook.Font                = 'Microsoft Sans Serif,10'

$AutoLogin                       = New-Object system.Windows.Forms.Button
$AutoLogin.text                  = "Setup Auto Login"
$AutoLogin.width                 = 175
$AutoLogin.height                = 25
$AutoLogin.location              = New-Object System.Drawing.Point(7,60)
$AutoLogin.Font                  = 'Microsoft Sans Serif,10'

$A2EExport                       = New-Object system.Windows.Forms.Button
$A2EExport.text                  = "Export License and Profile 1 Info"
$A2EExport.width                 = 220
$A2EExport.height                = 30
$A2EExport.location              = New-Object System.Drawing.Point(10,110)
$A2EExport.Font                  = 'Microsoft Sans Serif,10'

$GPOResult                       = New-Object system.Windows.Forms.Button
$GPOResult.text                  = "Get Group Policy Results"
$GPOResult.width                 = 175
$GPOResult.height                = 25
$GPOResult.location              = New-Object System.Drawing.Point(7,90)
$GPOResult.Font                  = 'Microsoft Sans Serif,10'

$ShellUpdate                     = New-Object system.Windows.Forms.Button
$ShellUpdate.text                = "Get MS Online Azure AD - Update"
$ShellUpdate.width               = 250
$ShellUpdate.height              = 30
$ShellUpdate.location            = New-Object System.Drawing.Point(10,30)
$ShellUpdate.Font                = 'Microsoft Sans Serif,10'

$OutlookAddinDisable             = New-Object system.Windows.Forms.Button
$OutlookAddinDisable.text        = "Disable Outlook Add-ins"
$OutlookAddinDisable.width       = 175
$OutlookAddinDisable.height      = 25
$OutlookAddinDisable.location    = New-Object System.Drawing.Point(7,120)
$OutlookAddinDisable.Font        = 'Microsoft Sans Serif,10'

$OutlookAddinsEnable             = New-Object system.Windows.Forms.Button
$OutlookAddinsEnable.text        = "Enable Outlook Add-ins"
$OutlookAddinsEnable.width       = 175
$OutlookAddinsEnable.height      = 25
$OutlookAddinsEnable.location    = New-Object System.Drawing.Point(7,150)
$OutlookAddinsEnable.Font        = 'Microsoft Sans Serif,10'

$Windows10Virgin                 = New-Object system.Windows.Forms.Button
$Windows10Virgin.text            = "Virginize Windows 10"
$Windows10Virgin.width           = 175
$Windows10Virgin.height          = 25
$Windows10Virgin.location        = New-Object System.Drawing.Point(7,180)
$Windows10Virgin.Font            = 'Microsoft Sans Serif,10'

$ExchangeSideLabel               = New-Object system.Windows.Forms.Label
$ExchangeSideLabel.text          = "Exchange Services"
$ExchangeSideLabel.AutoSize      = $true
$ExchangeSideLabel.width         = 25
$ExchangeSideLabel.height        = 10
$ExchangeSideLabel.location      = New-Object System.Drawing.Point(5,6)
$ExchangeSideLabel.Font          = 'Microsoft Sans Serif,10,style=Bold'

$A2ELable                        = New-Object system.Windows.Forms.Label
$A2ELable.text                   = "Add2Exchange Services"
$A2ELable.AutoSize               = $true
$A2ELable.width                  = 25
$A2ELable.height                 = 10
$A2ELable.location               = New-Object System.Drawing.Point(5,6)
$A2ELable.Font                   = 'Microsoft Sans Serif,10,style=Bold'

$WindowsServices                 = New-Object system.Windows.Forms.Label
$WindowsServices.text            = "Windows Services"
$WindowsServices.AutoSize        = $true
$WindowsServices.width           = 25
$WindowsServices.height          = 10
$WindowsServices.location        = New-Object System.Drawing.Point(5,6)
$WindowsServices.Font            = 'Microsoft Sans Serif,10,style=Bold'

$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 280
$Panel1.width                    = 270
$Panel1.location                 = New-Object System.Drawing.Point(307,279)

$A2ESQL                          = New-Object system.Windows.Forms.Label
$A2ESQL.text                     = "Add2Exchange SQL Services"
$A2ESQL.AutoSize                 = $true
$A2ESQL.width                    = 25
$A2ESQL.height                   = 10
$A2ESQL.location                 = New-Object System.Drawing.Point(5,6)
$A2ESQL.Font                     = 'Microsoft Sans Serif,10,style=Bold'

$SQLCommands                     = New-Object system.Windows.Forms.Button
$SQLCommands.text                = "SQL Coming soon!"
$SQLCommands.width               = 230
$SQLCommands.height              = 163
$SQLCommands.location            = New-Object System.Drawing.Point(22,55)
$SQLCommands.Font                = 'Microsoft Sans Serif,10'

$A2Eupdates                      = New-Object system.Windows.Forms.Button
$A2Eupdates.text                 = "Update Add2Exchange"
$A2Eupdates.width                = 220
$A2Eupdates.height               = 30
$A2Eupdates.location             = New-Object System.Drawing.Point(10,150)
$A2Eupdates.Font                 = 'Microsoft Sans Serif,10'

$DIBLOGO                         = New-Object system.Windows.Forms.PictureBox
$DIBLOGO.width                   = 220
$DIBLOGO.height                  = 45
$DIBLOGO.location                = New-Object System.Drawing.Point(8,10)
$DIBLOGO.imageLocation           = "\\Fileserv-db1\Work\Advantage International\Ai DIB Logos\DIBlogo\DidItBetterSoftware_logo.png"
$DIBLOGO.SizeMode                = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$A2ETools.controls.AddRange(@($ExchangePermissions,$Add2Exchange,$WindowsTasks,$Panel1,$DIBLOGO))
$ExchangePermissions.controls.AddRange(@($O365onprem,$ShellUpdate,$ExchangeSideLabel))
$WindowsTasks.controls.AddRange(@($UAC,$AutoLogin,$GPOResult,$OutlookAddinDisable,$OutlookAddinsEnable,$Windows10Virgin,$WindowsServices))
$Add2Exchange.controls.AddRange(@($Favorites,$Add2Outlook,$A2EExport,$A2ELable,$A2Eupdates))
$Panel1.controls.AddRange(@($A2ESQL,$SQLCommands))

#region gui events {

$O365onprem.Add_Click({  })
$ShellUpdate.Add_Click({PowershellUpdate})
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


############# PowerShell Updater #############

Function PowershellUpdate {
Write-Host "Chekcing version of powershell"
$PSVersionTable.PSVersion

$confirmation = Read-Host "Would you like me to update powershell? [Y/N]"
if ($confirmation -eq 'y') {
  

Write-Host "Updating Powershell"
Write-Host "Downloading the latest powershell Update"
$url = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/W2K12-KB3191565-x64.msu"
$output = "$PSScriptRoot\Powershell_5.msu"
$start_time = Get-Date


(New-Object System.Net.WebClient).DownloadFile($url, $output)

Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

}
$confirmation = Read-Host "Would you like me to Install MSonline Module? [Y/N]"
if ($confirmation -eq 'y') {
Write-Host "Adding Azure MSonline module"
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-Module MSonline -Confirm:$false -WarningAction "Inquire"

}
Write-Host "Ok...Done..."
}



# End Scripting


##############################################################

[void]$A2ETools.ShowDialog()