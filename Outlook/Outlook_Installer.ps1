<#
.NAME
    Outlook Installer
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Outlook365Installer             = New-Object system.Windows.Forms.Form
$Outlook365Installer.ClientSize  = '279,329'
$Outlook365Installer.text        = "Outlook 365 Install"
$Outlook365Installer.TopMost     = $false

$ProRetail                       = New-Object system.Windows.Forms.Label
$ProRetail.text                  = "Office 365 Pro Retail"
$ProRetail.AutoSize              = $true
$ProRetail.width                 = 25
$ProRetail.height                = 10
$ProRetail.location              = New-Object System.Drawing.Point(21,30)
$ProRetail.Font                  = 'Microsoft Sans Serif,14,style=Bold,Underline'

$ProBusiness                     = New-Object system.Windows.Forms.Label
$ProBusiness.text                = "Office 365 Business Retail"
$ProBusiness.AutoSize            = $true
$ProBusiness.width               = 25
$ProBusiness.height              = 10
$ProBusiness.location            = New-Object System.Drawing.Point(20,170)
$ProBusiness.Font                = 'Microsoft Sans Serif,14,style=Bold,Underline'

$Pro32                           = New-Object system.Windows.Forms.Button
$Pro32.text                      = "O365 Pro x86"
$Pro32.width                     = 200
$Pro32.height                    = 30
$Pro32.location                  = New-Object System.Drawing.Point(20,60)
$Pro32.Font                      = 'Microsoft Sans Serif,12'

$Pro64                           = New-Object system.Windows.Forms.Button
$Pro64.text                      = "O365 Pro x64"
$Pro64.width                     = 200
$Pro64.height                    = 30
$Pro64.location                  = New-Object System.Drawing.Point(20,100)
$Pro64.Font                      = 'Microsoft Sans Serif,12'

$Business32                      = New-Object system.Windows.Forms.Button
$Business32.text                 = "O365 Business x86"
$Business32.width                = 200
$Business32.height               = 30
$Business32.location             = New-Object System.Drawing.Point(20,200)
$Business32.Font                 = 'Microsoft Sans Serif,12'

$Business64                      = New-Object system.Windows.Forms.Button
$Business64.text                 = "O365 Business x64"
$Business64.width                = 200
$Business64.height               = 30
$Business64.location             = New-Object System.Drawing.Point(20,240)
$Business64.Font                 = 'Microsoft Sans Serif,12'

$Outlook365Installer.controls.AddRange(@($ProRetail,$ProBusiness,$Pro32,$Pro64,$Business32,$Business64))

$Pro32.Add_Click({Start-Process cmd -Argumentlist "/c Pro_Retailx86.cmd" -WorkingDirectory ".\Setup Files"})
$Pro64.Add_Click({Start-Process cmd -Argumentlist "/c Pro_Retailx64.cmd" -WorkingDirectory ".\Setup Files"})
$Business64.Add_Click({Start-Process cmd -Argumentlist "/c Busimess_Retailx64.cmd" -WorkingDirectory ".\Setup Files"})
$Business32.Add_Click({Start-Process cmd -Argumentlist "/c Busimess_Retailx86.cmd" -WorkingDirectory ".\Setup Files"})


[void]$Outlook365Installer.ShowDialog()

# End Scripting