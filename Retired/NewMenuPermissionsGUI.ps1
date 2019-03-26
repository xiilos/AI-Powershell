<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Add2Exchange Permissions
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$PermissionsonPremisandO365      = New-Object system.Windows.Forms.Form
$PermissionsonPremisandO365.ClientSize  = '523,530'
$PermissionsonPremisandO365.text  = "Add2Exchange Permissions"
$PermissionsonPremisandO365.TopMost  = $false

$ServiceAccountName              = New-Object system.Windows.Forms.TextBox
$ServiceAccountName.multiline    = $false
$ServiceAccountName.width        = 250
$ServiceAccountName.height       = 20
$ServiceAccountName.location     = New-Object System.Drawing.Point(14,105)
$ServiceAccountName.Font         = 'Microsoft Sans Serif,10'

$GO                              = New-Object system.Windows.Forms.Button
$GO.text                         = "GO"
$GO.width                        = 60
$GO.height                       = 30
$GO.location                     = New-Object System.Drawing.Point(368,449)
$GO.Font                         = 'Microsoft Sans Serif,10'

$Cancel                          = New-Object system.Windows.Forms.Button
$Cancel.text                     = "Cancel"
$Cancel.width                    = 60
$Cancel.height                   = 30
$Cancel.location                 = New-Object System.Drawing.Point(454,449)
$Cancel.Font                     = 'Microsoft Sans Serif,10'

$ExchangeServerName              = New-Object system.Windows.Forms.TextBox
$ExchangeServerName.multiline    = $false
$ExchangeServerName.width        = 250
$ExchangeServerName.height       = 20
$ExchangeServerName.location     = New-Object System.Drawing.Point(14,38)
$ExchangeServerName.Font         = 'Microsoft Sans Serif,10'

$Exchange2010                    = New-Object system.Windows.Forms.CheckBox
$Exchange2010.text               = "Exchange 2010"
$Exchange2010.AutoSize           = $false
$Exchange2010.width              = 105
$Exchange2010.height             = 15
$Exchange2010.location           = New-Object System.Drawing.Point(295,46)
$Exchange2010.Font               = 'Microsoft Sans Serif,10'

$Exchange2013_2016               = New-Object system.Windows.Forms.CheckBox
$Exchange2013_2016.text          = "Exchange 2013-2016"
$Exchange2013_2016.AutoSize      = $false
$Exchange2013_2016.width         = 143
$Exchange2013_2016.height        = 13
$Exchange2013_2016.location      = New-Object System.Drawing.Point(296,75)
$Exchange2013_2016.Font          = 'Microsoft Sans Serif,10'

$ExchangeServName                = New-Object system.Windows.Forms.Label
$ExchangeServName.text           = "Exchange Server Name"
$ExchangeServName.AutoSize       = $true
$ExchangeServName.width          = 25
$ExchangeServName.height         = 10
$ExchangeServName.location       = New-Object System.Drawing.Point(14,14)
$ExchangeServName.Font           = 'Microsoft Sans Serif,12,style=Bold,Underline'

$UerName_TenentAdmin             = New-Object system.Windows.Forms.Label
$UerName_TenentAdmin.text        = "Username or Tenent Admin"
$UerName_TenentAdmin.AutoSize    = $true
$UerName_TenentAdmin.width       = 25
$UerName_TenentAdmin.height      = 10
$UerName_TenentAdmin.location    = New-Object System.Drawing.Point(14,82)
$UerName_TenentAdmin.Font        = 'Microsoft Sans Serif,12,style=Bold,Underline'

$LogonMethod                     = New-Object system.Windows.Forms.Label
$LogonMethod.text                = "How are We Connecting?"
$LogonMethod.AutoSize            = $true
$LogonMethod.width               = 25
$LogonMethod.height              = 10
$LogonMethod.location            = New-Object System.Drawing.Point(295,12)
$LogonMethod.Font                = 'Microsoft Sans Serif,12,style=Bold,Underline'

$Office365                       = New-Object system.Windows.Forms.CheckBox
$Office365.text                  = "Office 365"
$Office365.AutoSize              = $false
$Office365.width                 = 95
$Office365.height                = 20
$Office365.location              = New-Object System.Drawing.Point(296,105)
$Office365.Font                  = 'Microsoft Sans Serif,10'

$ServiceAccount                  = New-Object system.Windows.Forms.Label
$ServiceAccount.text             = "Sync Service Account Name"
$ServiceAccount.AutoSize         = $true
$ServiceAccount.width            = 25
$ServiceAccount.height           = 10
$ServiceAccount.location         = New-Object System.Drawing.Point(14,147)
$ServiceAccount.Font             = 'Microsoft Sans Serif,12,style=Bold,Underline'

$SyncAccountName                 = New-Object system.Windows.Forms.TextBox
$SyncAccountName.multiline       = $false
$SyncAccountName.width           = 250
$SyncAccountName.height          = 20
$SyncAccountName.location        = New-Object System.Drawing.Point(14,173)
$SyncAccountName.Font            = 'Microsoft Sans Serif,10'

$AddPermissions                  = New-Object system.Windows.Forms.CheckBox
$AddPermissions.text             = "Add Add2Exchange Permissions to All Users"
$AddPermissions.AutoSize         = $false
$AddPermissions.width            = 350
$AddPermissions.height           = 20
$AddPermissions.location         = New-Object System.Drawing.Point(121,258)
$AddPermissions.Font             = 'Microsoft Sans Serif,10'

$PermissionsChoice               = New-Object system.Windows.Forms.Label
$PermissionsChoice.text          = "How would you Like to Proceed?"
$PermissionsChoice.AutoSize      = $true
$PermissionsChoice.width         = 25
$PermissionsChoice.height        = 10
$PermissionsChoice.location      = New-Object System.Drawing.Point(121,221)
$PermissionsChoice.Font          = 'Microsoft Sans Serif,12,style=Bold,Underline'

$RemovePermissions               = New-Object system.Windows.Forms.CheckBox
$RemovePermissions.text          = "Remove Add2Exchange Permissions from All Users"
$RemovePermissions.AutoSize      = $false
$RemovePermissions.width         = 350
$RemovePermissions.height        = 20
$RemovePermissions.location      = New-Object System.Drawing.Point(122,288)
$RemovePermissions.Font          = 'Microsoft Sans Serif,10'

$PermissionsonPremisandO365.controls.AddRange(@($ServiceAccountName,$GO,$Cancel,$ExchangeServerName,$Exchange2010,$Exchange2013_2016,$ExchangeServName,$UerName_TenentAdmin,
$LogonMethod,$Office365,$ServiceAccount,$SyncAccountName,$AddPermissions,$PermissionsChoice,$RemovePermissions))

$Cancel.Add_Click({  })
$GO.Add_Click({Write-Host "You picked $Exchange2010.checkstate and the server name is $ExchangeServerName.textbox"})



#Write your logic code here

[void]$PermissionsonPremisandO365.ShowDialog()