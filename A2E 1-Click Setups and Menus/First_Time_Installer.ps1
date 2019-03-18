<# 
.NAME
    Add2Exchange First Time Installer
#>

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$QucikStart_Menu                 = New-Object system.Windows.Forms.Form
$QucikStart_Menu.ClientSize      = '506,154'
$QucikStart_Menu.text            = "Add2Exchange Quick Start"
$QucikStart_Menu.TopMost         = $false

$whatdoesthisdo                  = New-Object system.Windows.Forms.ToolTip
$whatdoesthisdo.ToolTipTitle     = "whatdoesthisdo"
$whatdoesthisdo.isBalloon        = $true

$DIBLogo                         = New-Object system.Windows.Forms.PictureBox
$DIBLogo.width                   = 249
$DIBLogo.height                  = 76
$DIBLogo.location                = New-Object System.Drawing.Point(241,29)
$DIBLogo.imageLocation           = ".\Setup\DidItBetter_logo.png"
$DIBLogo.SizeMode                = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$FIrstTimeInstall                = New-Object system.Windows.Forms.Label
$FIrstTimeInstall.text           = "First Time Installer"
$FIrstTimeInstall.AutoSize       = $true
$FIrstTimeInstall.width          = 25
$FIrstTimeInstall.height         = 10
$FIrstTimeInstall.location       = New-Object System.Drawing.Point(16,27)
$FIrstTimeInstall.Font           = 'Microsoft Sans Serif,14,style=Underline'

$Step1                           = New-Object system.Windows.Forms.Button
$Step1.text                      = "Step 1: Run Me First"
$Step1.width                     = 165
$Step1.height                    = 30
$Step1.location                  = New-Object System.Drawing.Point(16,60)
$Step1.Font                      = 'Microsoft Sans Serif,10'

$Step2                           = New-Object system.Windows.Forms.Button
$Step2.text                      = "Step 2: Setup"
$Step2.width                     = 165
$Step2.height                    = 30
$Step2.location                  = New-Object System.Drawing.Point(16,105)
$Step2.Font                      = 'Microsoft Sans Serif,10'

$QucikStart_Menu.controls.AddRange(@($DIBLogo,$FIrstTimeInstall,$Step1,$Step2))

$Step1.Add_Click({.\Step1Initialize.ps1})
$Step2.Add_Click({.\Step2Setup.ps1})



#Functions Start Below Here

[void]$QucikStart_Menu.ShowDialog()