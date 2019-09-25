#Pathing

$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
    if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
        Write-Host "Secure Location Exists...Resuming"
    }
    Else {
        Write-Host "Creating Secure Location"
        New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
    }

    Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"

#Start Script

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Permissions_Menu   = New-Object system.Windows.Forms.Form
$Add2Exchange_Permissions_Menu.ClientSize  = '454,583'
$Add2Exchange_Permissions_Menu.text  = "DiditBetter Software Auto Permissions Setup"
$Add2Exchange_Permissions_Menu.BackColor  = "#ffffff"
$Add2Exchange_Permissions_Menu.TopMost  = $false

$DIB_Logo                        = New-Object system.Windows.Forms.PictureBox
$DIB_Logo.width                  = 149
$DIB_Logo.height                 = 50
$DIB_Logo.location               = New-Object System.Drawing.Point(14,527)
$DIB_Logo.imageLocation          = "./Diditbetter_logo.png"
$DIB_Logo.SizeMode               = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$ExchangeServerName_Label        = New-Object system.Windows.Forms.Label
$ExchangeServerName_Label.text   = "Exchange Server Name"
$ExchangeServerName_Label.AutoSize  = $true
$ExchangeServerName_Label.visible  = $true
$ExchangeServerName_Label.width  = 120
$ExchangeServerName_Label.height  = 10
$ExchangeServerName_Label.location  = New-Object System.Drawing.Point(15,25)
$ExchangeServerName_Label.Font   = 'Microsoft Sans Serif,10,style=Bold,Underline'

$ExServ_Name_txt                 = New-Object system.Windows.Forms.TextBox
$ExServ_Name_txt.multiline       = $false
$ExServ_Name_txt.width           = 200
$ExServ_Name_txt.height          = 20
$ExServ_Name_txt.location        = New-Object System.Drawing.Point(15,45)
$ExServ_Name_txt.Font            = 'Microsoft Sans Serif,10'

$Create_Task                     = New-Object system.Windows.Forms.Button
$Create_Task.text                = "Create Task!"
$Create_Task.width               = 140
$Create_Task.height              = 40
$Create_Task.location            = New-Object System.Drawing.Point(275,517)
$Create_Task.Font                = 'Microsoft Sans Serif,10,style=Bold'

$O365_GA_Label                   = New-Object system.Windows.Forms.Label
$O365_GA_Label.text              = "Global Admin Account Name"
$O365_GA_Label.AutoSize          = $true
$O365_GA_Label.width             = 25
$O365_GA_Label.height            = 10
$O365_GA_Label.location          = New-Object System.Drawing.Point(15,80)
$O365_GA_Label.Font              = 'Microsoft Sans Serif,10,style=Bold,Underline'

$GB_Admin_txt                    = New-Object system.Windows.Forms.TextBox
$GB_Admin_txt.multiline          = $false
$GB_Admin_txt.width              = 200
$GB_Admin_txt.height             = 20
$GB_Admin_txt.location           = New-Object System.Drawing.Point(15,100)
$GB_Admin_txt.Font               = 'Microsoft Sans Serif,10'

$Sync_Account_Label              = New-Object system.Windows.Forms.Label
$Sync_Account_Label.text         = "Sync Service Account Name"
$Sync_Account_Label.AutoSize     = $true
$Sync_Account_Label.width        = 25
$Sync_Account_Label.height       = 10
$Sync_Account_Label.location     = New-Object System.Drawing.Point(15,135)
$Sync_Account_Label.Font         = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Sync_Accoun_txt                 = New-Object system.Windows.Forms.TextBox
$Sync_Accoun_txt.multiline       = $false
$Sync_Accoun_txt.width           = 200
$Sync_Accoun_txt.height          = 20
$Sync_Accoun_txt.location        = New-Object System.Drawing.Point(15,155)
$Sync_Accoun_txt.Font            = 'Microsoft Sans Serif,10'

$O365_Check                      = New-Object system.Windows.Forms.CheckBox
$O365_Check.text                 = "Office 365"
$O365_Check.AutoSize             = $false
$O365_Check.width                = 175
$O365_Check.height               = 20
$O365_Check.location             = New-Object System.Drawing.Point(270,470)
$O365_Check.Font                 = 'Microsoft Sans Serif,10,style=Bold'

$On_Premise_Check                = New-Object system.Windows.Forms.CheckBox
$On_Premise_Check.text           = "Exchange On Premise"
$On_Premise_Check.AutoSize       = $false
$On_Premise_Check.width          = 175
$On_Premise_Check.height         = 20
$On_Premise_Check.location       = New-Object System.Drawing.Point(270,440)
$On_Premise_Check.Font           = 'Microsoft Sans Serif,10,style=Bold'

$All_Perm_Check                  = New-Object system.Windows.Forms.CheckBox
$All_Perm_Check.text             = "Give Permissions to Everyone"
$All_Perm_Check.AutoSize         = $false
$All_Perm_Check.width            = 225
$All_Perm_Check.height           = 20
$All_Perm_Check.location         = New-Object System.Drawing.Point(16,440)
$All_Perm_Check.Font             = 'Microsoft Sans Serif,10,style=Bold'

$Dist_List_Check                 = New-Object system.Windows.Forms.CheckBox
$Dist_List_Check.text            = "Only to Distribution List"
$Dist_List_Check.AutoSize        = $false
$Dist_List_Check.width           = 225
$Dist_List_Check.height          = 20
$Dist_List_Check.location        = New-Object System.Drawing.Point(16,470)
$Dist_List_Check.Font            = 'Microsoft Sans Serif,10,style=Bold'

$Dynamic_Check                   = New-Object system.Windows.Forms.CheckBox
$Dynamic_Check.text              = "Only to Dynamic Distribution List"
$Dynamic_Check.AutoSize          = $false
$Dynamic_Check.width             = 250
$Dynamic_Check.height            = 20
$Dynamic_Check.location          = New-Object System.Drawing.Point(16,500)
$Dynamic_Check.Font              = 'Microsoft Sans Serif,10,style=Bold'

$Dist_List_Label                 = New-Object system.Windows.Forms.Label
$Dist_List_Label.text            = "Distribution List Name"
$Dist_List_Label.AutoSize        = $true
$Dist_List_Label.width           = 25
$Dist_List_Label.height          = 10
$Dist_List_Label.location        = New-Object System.Drawing.Point(15,190)
$Dist_List_Label.Font            = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Dist_Name_txt                   = New-Object system.Windows.Forms.TextBox
$Dist_Name_txt.multiline         = $false
$Dist_Name_txt.width             = 200
$Dist_Name_txt.height            = 20
$Dist_Name_txt.location          = New-Object System.Drawing.Point(15,210)
$Dist_Name_txt.Font              = 'Microsoft Sans Serif,10'

$Dynamic_Label                   = New-Object system.Windows.Forms.Label
$Dynamic_Label.text              = "Dynamic Distribution List Name"
$Dynamic_Label.AutoSize          = $true
$Dynamic_Label.width             = 25
$Dynamic_Label.height            = 10
$Dynamic_Label.location          = New-Object System.Drawing.Point(15,245)
$Dynamic_Label.Font              = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Dynamix_txt                     = New-Object system.Windows.Forms.TextBox
$Dynamix_txt.multiline           = $false
$Dynamix_txt.width               = 200
$Dynamix_txt.height              = 20
$Dynamix_txt.location            = New-Object System.Drawing.Point(15,265)
$Dynamix_txt.Font                = 'Microsoft Sans Serif,10'

$Static_Label                    = New-Object system.Windows.Forms.Label
$Static_Label.text               = "Static Distribution List Name"
$Static_Label.AutoSize           = $true
$Static_Label.width              = 25
$Static_Label.height             = 10
$Static_Label.location           = New-Object System.Drawing.Point(15,300)
$Static_Label.Font               = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Static_txt                      = New-Object system.Windows.Forms.TextBox
$Static_txt.multiline            = $false
$Static_txt.width                = 200
$Static_txt.height               = 20
$Static_txt.location             = New-Object System.Drawing.Point(15,320)
$Static_txt.Font                 = 'Microsoft Sans Serif,10'

$Permissions_Options             = New-Object system.Windows.Forms.Label
$Permissions_Options.text        = "Select a single option"
$Permissions_Options.AutoSize    = $true
$Permissions_Options.width       = 25
$Permissions_Options.height      = 10
$Permissions_Options.location    = New-Object System.Drawing.Point(15,400)
$Permissions_Options.Font        = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Logon_Choice                    = New-Object system.Windows.Forms.Label
$Logon_Choice.text               = "Select logon method"
$Logon_Choice.AutoSize           = $true
$Logon_Choice.width              = 25
$Logon_Choice.height             = 10
$Logon_Choice.location           = New-Object System.Drawing.Point(270,400)
$Logon_Choice.Font               = 'Microsoft Sans Serif,10,style=Bold,Underline'

$Exchange_Admin_Password_Update   = New-Object system.Windows.Forms.Label
$Exchange_Admin_Password_Update.text  = "Updated:"
$Exchange_Admin_Password_Update.AutoSize  = $true
$Exchange_Admin_Password_Update.width  = 25
$Exchange_Admin_Password_Update.height  = 10
$Exchange_Admin_Password_Update.location  = New-Object System.Drawing.Point(250,70)
$Exchange_Admin_Password_Update.Font  = 'Microsoft Sans Serif,9'
$Exchange_Admin_Password_Update.ForeColor  = "#d0021b"

$Global_Admin_Password_Update    = New-Object system.Windows.Forms.Label
$Global_Admin_Password_Update.text  = "Updated:"
$Global_Admin_Password_Update.AutoSize  = $true
$Global_Admin_Password_Update.width  = 25
$Global_Admin_Password_Update.height  = 10
$Global_Admin_Password_Update.location  = New-Object System.Drawing.Point(250,140)
$Global_Admin_Password_Update.Font  = 'Microsoft Sans Serif,9'
$Global_Admin_Password_Update.ForeColor  = "#d0021b"

$UpdateCreds                     = New-Object system.Windows.Forms.Button
$UpdateCreds.text                = "Update Credentials"
$UpdateCreds.width               = 140
$UpdateCreds.height              = 40
$UpdateCreds.location            = New-Object System.Drawing.Point(275,320)
$UpdateCreds.Font                = 'Microsoft Sans Serif,10,style=Bold'

$EX_Pass                         = New-Object system.Windows.Forms.Button
$EX_Pass.text                    = "Exchange Admin Password"
$EX_Pass.width                   = 195
$EX_Pass.height                  = 30
$EX_Pass.location                = New-Object System.Drawing.Point(250,40)
$EX_Pass.Font                    = 'Microsoft Sans Serif,10,style=Bold,Italic'

$GB_Admin_Pass                   = New-Object system.Windows.Forms.Button
$GB_Admin_Pass.text              = "Global Admin Password"
$GB_Admin_Pass.width             = 195
$GB_Admin_Pass.height            = 30
$GB_Admin_Pass.location          = New-Object System.Drawing.Point(250,110)
$GB_Admin_Pass.Font              = 'Microsoft Sans Serif,10,style=Bold,Italic'

$Add2Exchange_Permissions_Menu.controls.AddRange(@($DIB_Logo,$ExchangeServerName_Label,$ExServ_Name_txt,$Create_Task,$O365_GA_Label,$GB_Admin_txt,$Sync_Account_Label,$Sync_Accoun_txt,$O365_Check,$On_Premise_Check,$All_Perm_Check,$Dist_List_Check,$Dynamic_Check,$Dist_List_Label,$Dist_Name_txt,$Dynamic_Label,$Dynamix_txt,$Static_Label,$Static_txt,$Permissions_Options,$Logon_Choice,$Exchange_Admin_Password_Update,$Global_Admin_Password_Update,$UpdateCreds,$EX_Pass,$GB_Admin_Pass))



$EX_Pass.Add_Click({Read-Host "Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file ".\Exchange_Server_Pass.txt"})
$GB_Admin_Pass.Add_Click({Read-Host "Global Admin Password" -assecurestring | convertfrom-securestring | out-file ".\GA_Admin_Pass.txt"})

$UpdateCreds.Add_Click({$ExServ_Name_txt.text | Out-File ".\Exchange_Server_Name.txt"})
$UpdateCreds.Add_Click({$GB_Admin_txt.text | Out-File ".\GA_Service_Account_Name.txt"})
$UpdateCreds.Add_Click({$Sync_Accoun_txt.text | Out-File ".\Sync_Account_Name.txt"})
$UpdateCreds.Add_Click({$Dist_Name_txt.text | Out-File ".\Dist_List_Name.txt"})
$UpdateCreds.Add_Click({$Dynamix_txt.text | Out-File ".\Dynamic_Name.txt"})
$UpdateCreds.Add_Click({$Static_txt.text | Out-File ".\Static_Name"})


$UpdateCreds.Add_Click({
    $wshell = New-Object -ComObject Wscript.Shell
                $answer = $wshell.Popup("Updated Credentials Successfully.", 0, "Permissions Creator", 0x1)
                if ($answer -eq 2) { break }
})

#Variables
$ExServ_Name_txt.text = get-content ".\Exchange_Server_Name.txt"
$GB_Admin_txt.text = get-content ".\GA_Service_Account_Name.txt"
$Sync_Accoun_txt.text = get-content ".\Sync_Account_Name.txt"
$Dist_Name_txt.text = get-content ".\Dist_List_Name.txt"
$Dynamix_txt.text = get-content ".\Dynamic_Name.txt"
$Static_txt.text = get-content ".\Static_Name"

$Exchange_Admin_Password_Update.text = Get-Item ".\Exchange_Server_Pass.txt" | ForEach-Object { $_.LastWriteTime }
$Global_Admin_Password_Update.text = Get-Item ".\GA_Admin_Pass.txt" | ForEach-Object { $_.LastWriteTime }

$Create_Task.Add_Click({  })

#Creating the Task

# Option 1: Office 365-Adding Permissions to All Users

$Create_Task.Add_Click({
    if ($O365_Check.Checked -and $All_Perm_Check.checked)
    { Write-Host "You chose to Add Permissions to All Users in Office 365"

    $Repeater = (New-TimeSpan -Minutes 360)
    $Duration = ([timeSpan]::maxvalue)
    $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
    Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "Add2Exchange Permissions" -Description "Adds Add2Exchange Permissions Automatically to All Users Mailboxes"
    Write-Host "Done"
}
 
})













[void]$Add2Exchange_Permissions_Menu.ShowDialog()

#End Script