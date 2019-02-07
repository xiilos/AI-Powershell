<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Add2Exchange Permissions
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$PermissionsonPremisandO365      = New-Object system.Windows.Forms.Form
$PermissionsonPremisandO365.ClientSize  = '572,530'
$PermissionsonPremisandO365.text  = "Add2Exchange Permissions"
$PermissionsonPremisandO365.TopMost  = $false

$LogonMethod                     = New-Object system.Windows.Forms.ComboBox
$LogonMethod.text                = "How would you like to connect?"
$LogonMethod.width               = 202
$LogonMethod.height              = 37
@('Exchange 2010','Exchange 2013','Exchange 2016','Office 365') | ForEach-Object {[void] $logonmethod.Items.Add($_)}
$LogonMethod.location            = New-Object System.Drawing.Point(19,111)
$LogonMethod.Font                = 'Microsoft Sans Serif,10'

$PermissionsChoice               = New-Object system.Windows.Forms.ComboBox
$PermissionsChoice.text          = "How do you want to add Permissions?"
$PermissionsChoice.width         = 500
$PermissionsChoice.height        = 28
@('Add Permissions to Everyone','Remove Permissions from Everyone','Add Permissions to a Distribution List','Remove Permissions from a Distribution List','Add Permissions to a Single User','Remove Permissions from a Single User') | ForEach-Object {[void] $PermissionsChoice.Items.Add($_)}
$PermissionsChoice.location      = New-Object System.Drawing.Point(19,169)
$PermissionsChoice.Font          = 'Microsoft Sans Serif,10'

$ServiceAccountName              = New-Object system.Windows.Forms.TextBox
$ServiceAccountName.multiline    = $false
$ServiceAccountName.text         = "Type in your Service Account Name"
$ServiceAccountName.width        = 217
$ServiceAccountName.height       = 20
$ServiceAccountName.location     = New-Object System.Drawing.Point(19,66)
$ServiceAccountName.Font         = 'Microsoft Sans Serif,10'

$DistListName                    = New-Object system.Windows.Forms.TextBox
$DistListName.multiline          = $false
$DistListName.text               = "What is the Name of your Distribution List?"
$DistListName.width              = 275
$DistListName.height             = 20
$DistListName.location           = New-Object System.Drawing.Point(21,224)
$DistListName.Font               = 'Microsoft Sans Serif,10'

$UserEmailAddress                = New-Object system.Windows.Forms.TextBox
$UserEmailAddress.multiline      = $false
$UserEmailAddress.text           = "Type in User Email Address"
$UserEmailAddress.width          = 270
$UserEmailAddress.height         = 20
$UserEmailAddress.location       = New-Object System.Drawing.Point(19,274)
$UserEmailAddress.Font           = 'Microsoft Sans Serif,10'

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
$ExchangeServerName.text         = "Exchange Server Name"
$ExchangeServerName.width        = 153
$ExchangeServerName.height       = 20
$ExchangeServerName.location     = New-Object System.Drawing.Point(21,16)
$ExchangeServerName.Font         = 'Microsoft Sans Serif,10'

$PermissionsonPremisandO365.controls.AddRange(@($LogonMethod,$PermissionsChoice,$ServiceAccountName,$DistListName,$UserEmailAddress,$GO,$Cancel,$ExchangeServerName))

#$GO.controls.AddRange(@($LogonMethod,$PermissionsChoice,$ServiceAccountName,$DistListName,$UserEmailAddress,$ExchangeServerName))

$GO.Add_click({Get-PSSession | Remove-PSSession Exit})

$cancel.Add_click({Get-PSSession | Remove-PSSession Exit})


#Write your logic code here

[void]$PermissionsonPremisandO365.ShowDialog()