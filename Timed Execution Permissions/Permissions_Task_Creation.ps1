<#
        .SYNOPSIS
        Permissions task creator

        .DESCRIPTION
        Updates EXO PS modules
        Saves and bit locks passwords and usernames for auto log in to exchange or office 365
        sets additional tasks for permissions to auto run using credentials provided

        .NOTES
        Version:        3.2023
        Author:         DidItBetter Software

    #>

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

#Pathing

$TestPath = "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
    
}
Else {
    
    New-Item -ItemType directory -Path "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"
}

#Check for MS Online Module
Write-Host "Checking for Exhange Online Module"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

IF (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Write-Host "Exchange Online Module Exists"

    $InstalledEXOv2 = ((Get-Module -Name ExchangeOnlineManagement -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()

    $LatestEXOv2 = (Find-Module -Name ExchangeOnlineManagement).Version.ToString()

    [PSCustomObject]@{
        Match = If ($InstalledEXOv2 -eq $LatestEXOv2) { Write-Host "You are on the latest Version" } 

        Else {
            Write-Host "Upgrading Modules..."
            Update-Module -Name ExchangeOnlineManagement -Force
            Write-Host "Success"
        }

    }


} 
Else {
    Write-Host "Module Does Not Exist"
    Write-Host "Downloading Exchange Online Management..."
    Install-Module –Name ExchangeOnlineManagement -Force
    Write-Host "Success"
}
        


#Start Script

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Add2Exchange_Permissions_Menu = New-Object system.Windows.Forms.Form
$Add2Exchange_Permissions_Menu.ClientSize = New-Object System.Drawing.Point(509, 718)
$Add2Exchange_Permissions_Menu.text = "DiditBetter Software Auto Permissions Setup"
$Add2Exchange_Permissions_Menu.TopMost = $false
$Add2Exchange_Permissions_Menu.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$ExchangeServerName_Label = New-Object system.Windows.Forms.Label
$ExchangeServerName_Label.text = "Exchange Server Name"
$ExchangeServerName_Label.AutoSize = $true
$ExchangeServerName_Label.visible = $true
$ExchangeServerName_Label.width = 120
$ExchangeServerName_Label.height = 10
$ExchangeServerName_Label.location = New-Object System.Drawing.Point(15, 14)
$ExchangeServerName_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$ExServ_Name_txt = New-Object system.Windows.Forms.TextBox
$ExServ_Name_txt.multiline = $false
$ExServ_Name_txt.text = "NetBIOS name"
$ExServ_Name_txt.width = 200
$ExServ_Name_txt.height = 20
$ExServ_Name_txt.location = New-Object System.Drawing.Point(15, 34)
$ExServ_Name_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ExServ_Name_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Create_Task = New-Object system.Windows.Forms.Button
$Create_Task.text = "Create Task!"
$Create_Task.width = 140
$Create_Task.height = 40
$Create_Task.location = New-Object System.Drawing.Point(350, 662)
$Create_Task.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Create_Task.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#417505")

$O365_GA_Label = New-Object system.Windows.Forms.Label
$O365_GA_Label.text = "O365 Exchange Admin Account Name"
$O365_GA_Label.AutoSize = $true
$O365_GA_Label.width = 25
$O365_GA_Label.height = 10
$O365_GA_Label.location = New-Object System.Drawing.Point(15, 130)
$O365_GA_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$GB_Admin_txt = New-Object system.Windows.Forms.TextBox
$GB_Admin_txt.multiline = $false
$GB_Admin_txt.text = "email address"
$GB_Admin_txt.width = 200
$GB_Admin_txt.height = 20
$GB_Admin_txt.location = New-Object System.Drawing.Point(15, 150)
$GB_Admin_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$GB_Admin_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Sync_Account_Label = New-Object system.Windows.Forms.Label
$Sync_Account_Label.text = "Sync Service Account Name"
$Sync_Account_Label.AutoSize = $true
$Sync_Account_Label.width = 25
$Sync_Account_Label.height = 10
$Sync_Account_Label.location = New-Object System.Drawing.Point(15, 185)
$Sync_Account_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Sync_Accoun_txt = New-Object system.Windows.Forms.TextBox
$Sync_Accoun_txt.multiline = $false
$Sync_Accoun_txt.text = "display name"
$Sync_Accoun_txt.width = 200
$Sync_Accoun_txt.height = 20
$Sync_Accoun_txt.location = New-Object System.Drawing.Point(15, 205)
$Sync_Accoun_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Sync_Accoun_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$O365_Check = New-Object system.Windows.Forms.CheckBox
$O365_Check.text = "Office 365"
$O365_Check.AutoSize = $false
$O365_Check.width = 175
$O365_Check.height = 20
$O365_Check.location = New-Object System.Drawing.Point(313, 589)
$O365_Check.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$On_Premise_Check = New-Object system.Windows.Forms.CheckBox
$On_Premise_Check.text = "Exchange On Premise"
$On_Premise_Check.AutoSize = $false
$On_Premise_Check.width = 175
$On_Premise_Check.height = 20
$On_Premise_Check.location = New-Object System.Drawing.Point(313, 569)
$On_Premise_Check.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$All_Perm_Check = New-Object system.Windows.Forms.CheckBox
$All_Perm_Check.text = "Give Permissions to Everyone"
$All_Perm_Check.AutoSize = $false
$All_Perm_Check.width = 225
$All_Perm_Check.height = 20
$All_Perm_Check.location = New-Object System.Drawing.Point(20, 569)
$All_Perm_Check.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Dist_List_Check = New-Object system.Windows.Forms.CheckBox
$Dist_List_Check.text = "Only to Distribution List"
$Dist_List_Check.AutoSize = $false
$Dist_List_Check.width = 225
$Dist_List_Check.height = 20
$Dist_List_Check.location = New-Object System.Drawing.Point(20, 589)
$Dist_List_Check.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Dynamic_Check = New-Object system.Windows.Forms.CheckBox
$Dynamic_Check.text = "Only to Dynamic Distribution List"
$Dynamic_Check.AutoSize = $false
$Dynamic_Check.width = 250
$Dynamic_Check.height = 20
$Dynamic_Check.location = New-Object System.Drawing.Point(20, 609)
$Dynamic_Check.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Dist_List_Label = New-Object system.Windows.Forms.Label
$Dist_List_Label.text = "Distribution List Name"
$Dist_List_Label.AutoSize = $true
$Dist_List_Label.width = 25
$Dist_List_Label.height = 10
$Dist_List_Label.location = New-Object System.Drawing.Point(15, 234)
$Dist_List_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Dist_Name_txt = New-Object system.Windows.Forms.TextBox
$Dist_Name_txt.multiline = $true
$Dist_Name_txt.text = "display name"
$Dist_Name_txt.width = 200
$Dist_Name_txt.height = 150
$Dist_Name_txt.location = New-Object System.Drawing.Point(15, 254)
$Dist_Name_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Dist_Name_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Dynamic_Label = New-Object system.Windows.Forms.Label
$Dynamic_Label.text = "Dynamic Distribution List Name"
$Dynamic_Label.AutoSize = $true
$Dynamic_Label.width = 25
$Dynamic_Label.height = 10
$Dynamic_Label.location = New-Object System.Drawing.Point(15, 430)
$Dynamic_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Dynamic_txt = New-Object system.Windows.Forms.TextBox
$Dynamic_txt.multiline = $false
$Dynamic_txt.text = "display name"
$Dynamic_txt.width = 200
$Dynamic_txt.height = 20
$Dynamic_txt.location = New-Object System.Drawing.Point(15, 450)
$Dynamic_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Dynamic_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Static_Label = New-Object system.Windows.Forms.Label
$Static_Label.text = "Static Distribution List Name"
$Static_Label.AutoSize = $true
$Static_Label.width = 25
$Static_Label.height = 10
$Static_Label.location = New-Object System.Drawing.Point(15, 485)
$Static_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Static_txt = New-Object system.Windows.Forms.TextBox
$Static_txt.multiline = $false
$Static_txt.text = "display name"
$Static_txt.width = 200
$Static_txt.height = 20
$Static_txt.location = New-Object System.Drawing.Point(15, 505)
$Static_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$Static_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Permissions_Options = New-Object system.Windows.Forms.Label
$Permissions_Options.text = "Select a single option"
$Permissions_Options.AutoSize = $true
$Permissions_Options.width = 25
$Permissions_Options.height = 10
$Permissions_Options.location = New-Object System.Drawing.Point(20, 544)
$Permissions_Options.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Logon_Choice = New-Object system.Windows.Forms.Label
$Logon_Choice.text = "Select logon method"
$Logon_Choice.AutoSize = $true
$Logon_Choice.width = 25
$Logon_Choice.height = 10
$Logon_Choice.location = New-Object System.Drawing.Point(309, 544)
$Logon_Choice.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Exchange_Admin_Password_Update = New-Object system.Windows.Forms.Label
$Exchange_Admin_Password_Update.text = "Updated:"
$Exchange_Admin_Password_Update.AutoSize = $true
$Exchange_Admin_Password_Update.width = 25
$Exchange_Admin_Password_Update.height = 10
$Exchange_Admin_Password_Update.location = New-Object System.Drawing.Point(270, 123)
$Exchange_Admin_Password_Update.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$Exchange_Admin_Password_Update.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$Global_Admin_Password_Update = New-Object system.Windows.Forms.Label
$Global_Admin_Password_Update.text = "Updated:"
$Global_Admin_Password_Update.AutoSize = $true
$Global_Admin_Password_Update.width = 25
$Global_Admin_Password_Update.height = 10
$Global_Admin_Password_Update.location = New-Object System.Drawing.Point(270, 181)
$Global_Admin_Password_Update.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$Global_Admin_Password_Update.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$UpdateCreds = New-Object system.Windows.Forms.Button
$UpdateCreds.text = "Update Credentials"
$UpdateCreds.width = 140
$UpdateCreds.height = 40
$UpdateCreds.location = New-Object System.Drawing.Point(270, 390)
$UpdateCreds.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$UpdateCreds.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#417505")

$EX_Pass = New-Object system.Windows.Forms.Button
$EX_Pass.text = "Exchange Admin Password"
$EX_Pass.width = 195
$EX_Pass.height = 30
$EX_Pass.location = New-Object System.Drawing.Point(270, 87)
$EX_Pass.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))

$GB_Admin_Pass = New-Object system.Windows.Forms.Button
$GB_Admin_Pass.text = "O365 Admin Password"
$GB_Admin_Pass.width = 195
$GB_Admin_Pass.height = 30
$GB_Admin_Pass.location = New-Object System.Drawing.Point(270, 144)
$GB_Admin_Pass.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))

$DualPermissions = New-Object system.Windows.Forms.CheckBox
$DualPermissions.text = "Dual Permissions"
$DualPermissions.AutoSize = $false
$DualPermissions.width = 175
$DualPermissions.height = 20
$DualPermissions.location = New-Object System.Drawing.Point(313, 609)
$DualPermissions.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$ExchangeAdmin = New-Object system.Windows.Forms.Label
$ExchangeAdmin.text = "Exchange Admin Account Name"
$ExchangeAdmin.AutoSize = $true
$ExchangeAdmin.width = 25
$ExchangeAdmin.height = 10
$ExchangeAdmin.location = New-Object System.Drawing.Point(15, 68)
$ExchangeAdmin.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$ExAdmin_txt = New-Object system.Windows.Forms.TextBox
$ExAdmin_txt.multiline = $false
$ExAdmin_txt.text = "domain\username"
$ExAdmin_txt.width = 200
$ExAdmin_txt.height = 20
$ExAdmin_txt.location = New-Object System.Drawing.Point(15, 93)
$ExAdmin_txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ExAdmin_txt.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Shell_Permissions = New-Object system.Windows.Forms.Button
$Shell_Permissions.text = "Auto Shell Permissions"
$Shell_Permissions.width = 212
$Shell_Permissions.height = 47
$Shell_Permissions.location = New-Object System.Drawing.Point(270, 14)
$Shell_Permissions.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$Shell_Permissions.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#195eae")

$LocalAccount = New-Object system.Windows.Forms.Label
$LocalAccount.text = "Local Account"
$LocalAccount.AutoSize = $true
$LocalAccount.width = 25
$LocalAccount.height = 10
$LocalAccount.location = New-Object System.Drawing.Point(270, 240)
$LocalAccount.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))

$Localacctname = New-Object system.Windows.Forms.TextBox
$Localacctname.multiline = $false
$Localacctname.width = 225
$Localacctname.height = 20
$Localacctname.location = New-Object System.Drawing.Point(270, 260)
$Localacctname.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$LocalPassword = New-Object system.Windows.Forms.Button
$LocalPassword.text = "Local Account Password"
$LocalPassword.width = 225
$LocalPassword.height = 30
$LocalPassword.location = New-Object System.Drawing.Point(270, 305)
$LocalPassword.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))

$LocalPassUpdate = New-Object system.Windows.Forms.Label
$LocalPassUpdate.text = "Updated:"
$LocalPassUpdate.AutoSize = $true
$LocalPassUpdate.width = 25
$LocalPassUpdate.height = 10
$LocalPassUpdate.location = New-Object System.Drawing.Point(270, 342)
$LocalPassUpdate.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$LocalPassUpdate.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$Add2Exchange_Permissions_Menu.controls.AddRange(@($DIB_Logo, $ExchangeServerName_Label, $ExServ_Name_txt, $Create_Task, $O365_GA_Label, $GB_Admin_txt, $Sync_Account_Label, $Sync_Accoun_txt, $O365_Check,
        $On_Premise_Check, $All_Perm_Check, $Dist_List_Check, $Dynamic_Check, $Dist_List_Label, $Dist_Name_txt, $Dynamic_Label, $Dynamic_txt, $Static_Label, $Static_txt, $Permissions_Options, $Logon_Choice,
        $Exchange_Admin_Password_Update, $Global_Admin_Password_Update, $UpdateCreds, $EX_Pass, $GB_Admin_Pass, $DualPermissions, $ExchangeAdmin, $ExAdmin_txt, $Shell_Permissions, $LocalAccount, $Localacctname,
        $LocalPassword, $LocalPassUpdate))

#Shell into Exchange or O365 Auto with Options
$Shell_Permissions.Add_Click( { 
        Push-Location "C:\Program Files (x86)\OpenDoor Software®\Add2Exchange\Setup"
        Start-Process Powershell .\Shell_Permissions.ps1 })


#Update Credentials Variables
Push-Location "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds"

#Rename Old Files
Rename-Item -Path ".\ServerUser.txt" -NewName "Exchange_Server_Admin.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\ServerPass.txt" -NewName "Exchange_Server_Pass.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\ExchangeName.txt" -NewName "Exchange_Server_Name.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\Dynamic_DistributionName.txt" -NewName "Dynamic_Name.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\Static_DistributionName.txt" -NewName "Static_Name.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\ServiceAccount.txt" -NewName "Sync_Account_Name.txt" -ErrorAction SilentlyContinue
Rename-Item -Path ".\DistributionName.txt" -NewName "Dist_List_Name.txt" -ErrorAction SilentlyContinue

#Text Field Get Content
[System.Security.Principal.WindowsIdentity]::GetCurrent().Name | Out-File ".\Local_Account_Name.txt"

$ExServ_Name_txt.text = get-content ".\Exchange_Server_Name.txt" -ErrorAction SilentlyContinue
$ExAdmin_txt.text = get-content ".\Exchange_Server_Admin.txt" -ErrorAction SilentlyContinue
$GB_Admin_txt.text = get-content ".\GA_Service_Account_Name.txt" -ErrorAction SilentlyContinue
$Sync_Accoun_txt.text = get-content ".\Sync_Account_Name.txt" -ErrorAction SilentlyContinue
$Dist_Name_txt.text = get-content -Raw ".\Dist_List_Name.txt" -ErrorAction SilentlyContinue
$Dynamic_txt.text = get-content ".\Dynamic_Name.txt" -ErrorAction SilentlyContinue
$Static_txt.text = get-content ".\Static_Name.txt" -ErrorAction SilentlyContinue
$Localacctname.text = get-content ".\Local_Account_Name.txt" -ErrorAction SilentlyContinue

#Password Update Time
$Exchange_Admin_Password_Update.text = Get-Item ".\Exchange_Server_Pass.txt"  -ErrorAction SilentlyContinue | ForEach-Object { $_.LastWriteTime }
$Global_Admin_Password_Update.text = Get-Item ".\GA_Admin_Pass.txt" -ErrorAction SilentlyContinue | ForEach-Object { $_.LastWriteTime }
$LocalPassUpdate.text = Get-Item ".\Local_Account_Pass.txt" -ErrorAction SilentlyContinue | ForEach-Object { $_.LastWriteTime }

#Password Input Give
$EX_Pass.Add_Click( { Read-Host "Exchange Admin Password" -assecurestring | convertfrom-securestring | out-file "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Pass.txt" })
$GB_Admin_Pass.Add_Click( { Read-Host "Office 365 Admin Password" -assecurestring | convertfrom-securestring | out-file "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Admin_Pass.txt" })
$LocalPassword.Add_Click( { Read-Host "Local Account Password" -assecurestring | convertfrom-securestring | out-file "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" })

#Click to Open Dist. Lists Text
$Dist_List_Label.Add_Click( { Start-Process "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dist_List_Name.txt" })

#Update Credentials
$UpdateCreds.Add_Click( { 
    
        $ExServ_Name_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Name.txt"
        $ExAdmin_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Exchange_Server_Admin.txt"
        $GB_Admin_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\GA_Service_Account_Name.txt"
        $Sync_Accoun_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Sync_Account_Name.txt"
        $Dist_Name_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dist_List_Name.txt" -encoding utf8
        $Dynamic_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Dynamic_Name.txt"
        $Static_txt.text | Out-File "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Static_Name.txt"

        $wshell = New-Object -ComObject Wscript.Shell
        $answer = $wshell.Popup("Updated Credentials Successfully.", 0, "Permissions Creator", 0x1)
        if ($answer -eq 2) { [environment]::exit(0) }
    })


# Select Option Blank Outs
$All_Perm_Check.Add_CheckStateChanged( {
        if ($All_Perm_Check.checked) { $Dist_List_Check.Enabled = $false }
        else { $Dist_List_Check.Enabled = $true }

        if ($All_Perm_Check.checked) { $Dynamic_Check.Enabled = $false }
        else { $Dynamic_Check.Enabled = $true }
    })

$Dist_List_Check.Add_CheckStateChanged( {
        if ($Dist_List_Check.checked) { $All_Perm_Check.Enabled = $false }
        else { $All_Perm_Check.Enabled = $true }

        if ($Dist_List_Check.checked) { $Dynamic_Check.Enabled = $false }
        else { $Dynamic_Check.Enabled = $true }
    })

$Dynamic_Check.Add_CheckStateChanged( {
        if ($Dynamic_Check.checked) { $All_Perm_Check.Enabled = $false }
        else { $All_Perm_Check.Enabled = $true }
    
        if ($Dynamic_Check.checked) { $Dist_List_Check.Enabled = $false }
        else { $Dist_List_Check.Enabled = $true }
    })

$O365_Check.Add_CheckStateChanged( {
        if ($O365_Check.checked) { $Dynamic_Check.Enabled = $false }
        else { $Dynamic_Check.Enabled = $true }
    
    })

$Dynamic_Check.Add_CheckStateChanged( {
        if ($Dynamic_Check.checked) { $O365_Check.Enabled = $false }
        else { $O365_Check.Enabled = $true }
    
    })

$O365_Check.Add_CheckStateChanged( {
        if ($O365_Check.checked) { $On_Premise_Check.Enabled = $false }
        else { $On_Premise_Check.Enabled = $true }
    })

$O365_Check.Add_CheckStateChanged( {
        if ($O365_Check.checked) { $DualPermissions.Enabled = $false }
        else { $DualPermissions.Enabled = $true }
    })

$On_Premise_Check.Add_CheckStateChanged( {
        if ($On_Premise_Check.checked) { $O365_Check.Enabled = $false }
        else { $O365_Check.Enabled = $true }
    })

$On_Premise_Check.Add_CheckStateChanged( {
        if ($On_Premise_Check.checked) { $DualPermissions.Enabled = $false }
        else { $DualPermissions.Enabled = $true }
    })

$DualPermissions.Add_CheckStateChanged( {
        if ($DualPermissions.checked) { $O365_Check.Enabled = $false }
        else { $O365_Check.Enabled = $true }

        if ($DualPermissions.checked) { $On_Premise_Check.Enabled = $false }
        else { $On_Premise_Check.Enabled = $true }
    })

$DualPermissions.Add_CheckStateChanged( {
        if ($DualPermissions.checked) { $Dynamic_Check.Enabled = $false }
        else { $Dynamic_Check.Enabled = $true }
    })


#Check for Powershell File Paths
$Location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\OpenDoor Software®\Add2Exchange' -Name "InstallLocation").InstallLocation
Set-Location $Location


#Creating the Task
# Option 1: Office 365-Adding Permissions to All Users
$Create_Task.Add_Click( {
    
        if ($O365_Check.checked -and $All_Perm_Check.checked) {
            Write-Host "You chose to Add Permissions to All Users in Office 365"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Permissions to All O365" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to All O365" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to All O365" -Description "Add Permissions to All Users in Office 365" -User $UserID -Password $Password
            Write-Host "Done"
            
        }


        # Option 2: Office 365-Add Permissions to a Distribution List in Office 365
        if ($O365_Check.checked -and $Dist_List_Check.checked) {
            Write-Host "You chose to Add Permissions to a Distribution List in Office 365"
            
            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Permissions to Distribution List O365" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to Distribution List O365" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dist_List_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to Distribution List O365" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List" -User $UserID -Password $Password
            Write-Host "Done" 
            
        }


        # Option 3: Office 365-Add Permissions to a Dynamic Distribution List in Office 365
        if ($O365_Check.checked -and $Dynamic_Check.checked) {
            Write-Host "You chose to Add Permissions to a Dynamic Distribution List in Office 365"
           
            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Permissions to Dynamic List O365" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to Dynamic List O365" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 720)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dynamic_Distribution.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to Dynamic List O365" -Description "Exports a Dynamic Distribution List of Users into a Static List of Users" -User $UserID -Password $Password
            Write-Host "Done"
            
        }


        # Option 4: Office 365 & On-Premise-Add Permissions to All
        if ($DualPermissions.checked -and $All_Perm_Check.checked) {
            Write-Host "You chose to Add Permissions to all users on Premise & Office 365"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Dual O365 Permissions to All" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Dual O365 Permissions to All" -Confirm:$false
            }
        
            if (Get-ScheduledTask "A2E Dual On-Premise Permissions to All" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "A2E Dual On-Premise Permissions to All" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_All_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Dual O365 Permissions to All" -Description "Add Permissions to All Users in Office 365 and On-Premise" -User $UserID -Password $Password
            Write-Host "Done"


            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_All_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Dual On-Premise Permissions to All" -Description "Add Permissions to All Users in Office 365 and On-Premise" -User $UserID -Password $Password
            Write-Host "Done"

        }


        # Option 5: Office 365 & On-Premis-Add Permissions to a Distribution List
        if ($DualPermissions.checked -and $Dist_List_Check.checked) {
            Write-Host "You chose to Add Permissions to a Distribution List in Office 365 and On-Premise"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Dual O365 Permissions to a Distribution List" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Dual O365 Permissions to a Distribution List" -Confirm:$false
            }
        
            if (Get-ScheduledTask "A2E Dual On-Premise Permissions to a Distribution List" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "A2E Dual On-Premise Permissions to a Distribution List" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\Office365_Dist_List_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Dual O365 Permissions to a Distribution List" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List" -User $UserID -Password $Password
            Write-Host "Done" 



            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_Dist_List_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Dual On-Premise Permissions to a Distribution List" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List" -User $UserID -Password $Password
            Write-Host "Done"
        }


        # Option 6: Exchange On Premise-Adding Permissions to All Users
        if ($On_Premise_Check.checked -and $All_Perm_Check.checked) {
            Write-Host "You chose to Add Permissions to All Users on Premise"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Permissions to All on Premise Exchange" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to All on Premise Exchange" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_All_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to All on Premise Exchange" -Description "Add Permissions to All Users on Premise Exchange" -User $UserID -Password $Password
            Write-Host "Done"
            
        }


        # Option 7: Exchange On Premise-Add Permissions to a Distribution List
        if ($On_Premise_Check.checked -and $Dist_List_Check.checked) {
            Write-Host "You chose to Add Permissions to a Distribution List on Premise"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }

            if (Get-ScheduledTask "A2E Permissions to Distribution List on Premise Exchange" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to Distribution List on Premise Exchange" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_Dist_List_Permissions.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to Distribution List on Premise Exchange" -Description "Adds Add2Exchange Permissions Automatically to a Distribution List" -User $UserID -Password $Password
            Write-Host "Done" 
            
        }


        # Option 8: Exchange On Premise-Add Permissions to a Dynamic Distribution List
        if ($On_Premise_Check.checked -and $Dynamic_Check.checked) {
            Write-Host "You chose to Add Permissions to a Dynamic Distribution List on Premise"

            #Check If old Tasks Already Exist
            if (Get-ScheduledTask "Add2Exchange Permissions" -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName "Add2Exchange Permissions" -Confirm:$false
            }
            
            if (Get-ScheduledTask "A2E Permissions to Dynamic Distribution List on Premise Exchange" -ErrorAction SilentlyContinue) {
                Write-Host "Add2Exchange Permissions Task Already Exists... Updating..."
                Unregister-ScheduledTask -TaskName "A2E Permissions to Dynamic Distribution List on Premise Exchange" -Confirm:$false
            }

            $Repeater = (New-TimeSpan -Minutes 360)
            $Duration = ([timeSpan]::maxvalue)
            $Trigger = New-JobTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval $Repeater -RepetitionDuration $Duration
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -WorkingDirectory $Location -Argument '-NoProfile -WindowStyle Hidden -Executionpolicy Bypass -file ".\Setup\Timed Permissions\2010-2019_Dynamic_Distribution.ps1"'
            $UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
            Register-ScheduledTask -Action $Action -RunLevel Highest -Trigger $Trigger -TaskName "A2E Permissions to Dynamic Distribution List on Premise Exchange" -Description "Adds Add2Exchange Permissions Automatically to a Dynamic Distribution List" -User $UserID -Password $Password
            Write-Host "Done" 
            
        }


    })


[void]$Add2Exchange_Permissions_Menu.ShowDialog()

#End Script