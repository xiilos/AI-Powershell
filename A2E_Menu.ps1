if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted


# Script #
<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    A2EMenu
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Menu               = New-Object system.Windows.Forms.Form
$Add2Exchange_Menu.ClientSize    = '956,450'
$Add2Exchange_Menu.text          = "Form"
$Add2Exchange_Menu.TopMost       = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "button"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(111,158)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$DIB_Logo                        = New-Object system.Windows.Forms.PictureBox
$DIB_Logo.width                  = 106
$DIB_Logo.height                 = 42
$DIB_Logo.location               = New-Object System.Drawing.Point(11,11)
$DIB_Logo.imageLocation          = "D:\download\Diditbetter_logo_final.png"
$DIB_Logo.SizeMode               = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Add2Exchange_Menu.controls.AddRange(@($Button1,$DIB_Logo))

$Button1.Add_Click({  })



# End Scripting