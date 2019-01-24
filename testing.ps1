<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    testing
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1177,563'
$Form.text                       = "Form"
$Form.TopMost                    = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "what does this do?"
$Button1.width                   = 323
$Button1.height                  = 105
$Button1.location                = New-Object System.Drawing.Point(119,61)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$whatdoesthisdo                  = New-Object system.Windows.Forms.ToolTip
$whatdoesthisdo.ToolTipTitle     = "whatdoesthisdo"
$whatdoesthisdo.isBalloon        = $false

$whatdoesthisdo.SetToolTip($Button1,'this runs through permissions when you click it')
$Form.controls.AddRange(@($Button1))

$Button1.Add_Click({  })



#Write your logic code here

[void]$Form.ShowDialog()