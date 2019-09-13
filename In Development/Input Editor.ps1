<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    A2EMenu
.DESCRIPTION
    Add2Exchange Menu
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Menu               = New-Object system.Windows.Forms.Form
$Add2Exchange_Menu.ClientSize    = '400,499'
$Add2Exchange_Menu.text          = "DiditBetter Software Add2Exchange Setup"
$Add2Exchange_Menu.BackColor     = "#ffffff"
$Add2Exchange_Menu.TopMost       = $false

$DIB_Logo                        = New-Object system.Windows.Forms.PictureBox
$DIB_Logo.width                  = 274
$DIB_Logo.height                 = 64
$DIB_Logo.location               = New-Object System.Drawing.Point(15,391)
$DIB_Logo.imageLocation          = "https://www.diditbetter.com/wp-content/uploads/2017/03/Diditbetter_logo_final.svg"
$DIB_Logo.SizeMode               = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$Label_Tools                     = New-Object system.Windows.Forms.Label
$Label_Tools.text                = "Exchange name"
$Label_Tools.AutoSize            = $true
$Label_Tools.visible             = $true
$Label_Tools.width               = 120
$Label_Tools.height              = 10
$Label_Tools.location            = New-Object System.Drawing.Point(18,37)
$Label_Tools.Font                = 'Microsoft Sans Serif,10,style=Bold,Underline'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 287
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(13,65)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Exchange Password"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(17,127)
$Label1.Font                     = 'Microsoft Sans Serif,10,style=Bold'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Submit!"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(262,311)
$Button1.Font                    = 'Microsoft Sans Serif,10'


$Add2Exchange_Menu.controls.AddRange(@($DIB_Logo,$Label_Tools,$TextBox1,$Label1,$Button1))



$Button1.Add_Click({$TextBox1.text | out-file "d:\downloads\ExchangeName.txt"})
$Button1.Add_Click({$TextBox1.text | out-file "d:\downloads\ExchangeName.txt"})

$TextBox1.text = get-content "d:\downloads\ExchangeName.txt"



#Write your logic code here



[void]$Add2Exchange_Menu.ShowDialog()