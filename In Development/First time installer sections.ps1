#Step 1: Account Creation
#Step 2: Upgrade .Net and Powershell if needed
#Step 3: Create zLibrary and Create Shortcuts
#Step 4: Install Outlook and Setup Profile
#Step 5: Mailbox Creation
#Step 6: Create a Mail Profile
#Step 7: Add Permissions (moved to step 11a)
#Step 8: Add Public Folder Permissions
#Step 9: Enable AutoLogon
#Step 10: Install Add2Exchange
#Step 11: Add Registry Favs
#Step 11a: Setup Timed Permissions
#Step 12: Cleanup




#Step 1: Account Creation

# Create a GUI window
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create User Account"
$form.Size = New-Object System.Drawing.Size(300,250)
$form.StartPosition = "CenterScreen"

# Create labels for user input fields
$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Location = New-Object System.Drawing.Point(10,20)
$lbl1.Size = New-Object System.Drawing.Size(120,20)
$lbl1.Text = "Username:"
$form.Controls.Add($lbl1)

$lbl2 = New-Object System.Windows.Forms.Label
$lbl2.Location = New-Object System.Drawing.Point(10,60)
$lbl2.Size = New-Object System.Drawing.Size(120,20)
$lbl2.Text = "Password:"
$form.Controls.Add($lbl2)

# Create text boxes for user input
$txt1 = New-Object System.Windows.Forms.TextBox
$txt1.Location = New-Object System.Drawing.Point(130,20)
$txt1.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($txt1)

$txt2 = New-Object System.Windows.Forms.TextBox
$txt2.Location = New-Object System.Drawing.Point(130,60)
$txt2.Size = New-Object System.Drawing.Size(150,20)
$txt2.PasswordChar = '*'
$form.Controls.Add($txt2)

# Create a button to submit user input and create the account
$btn1 = New-Object System.Windows.Forms.Button
$btn1.Location = New-Object System.Drawing.Point(100,100)
$btn1.Size = New-Object System.Drawing.Size(100,25)
$btn1.Text = "Create Account"
$btn1.Add_Click({
    $username = $txt1.Text
    $password = $txt2.Text
    $computername = Read-Host "Enter the computer name"
    $user = [ADSI]"WinNT://$computername"
    $newUser = $user.Create("User", $username)
    $newUser.SetPassword($password)
    $newUser.SetInfo()
    $newUser.FullName = $username
    $newUser.SetInfo()
    $newUser.UserFlags = 0x10000
    $newUser.SetInfo()
    $path = "C:\user_account.txt"
    $content = "Username: $username`nPassword: $password`nComputer: $computername"
    Set-Content -Path $path -Value $content
    $form.Close()
})
$form.Controls.Add($btn1)

# Display the GUI window
$form.ShowDialog()



#Step 2: Upgrade .Net and Powershell if needed

# Check if .NET Framework 4.8 is installed
$net48 = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
if (!$net48) {
    # Install .NET Framework 4.8
    Invoke-WebRequest 'https://download.microsoft.com/download/D/9/4/D948E57B-7B76-4E22-8E44-F6D84ED03A2C/NDP48-KB4503541-x86-x64-AllOS-ENU.exe' -OutFile 'NDP48-KB4503541-x86-x64-AllOS-ENU.exe'
    Start-Process -FilePath 'NDP48-KB4503541-x86-x64-AllOS-ENU.exe' -ArgumentList '/quiet /norestart' -Wait
    Remove-Item 'NDP48-KB4503541-x86-x64-AllOS-ENU.exe'
    Write-Host '.NET Framework 4.8 has been installed.'
}

# Check if PowerShell 7.x is installed
$pwsh = Get-Command 'pwsh' -ErrorAction SilentlyContinue
if (!$pwsh) {
    # Install PowerShell 7.x
    Invoke-WebRequest 'https://github.com/PowerShell/PowerShell/releases/download/v7.1.3/PowerShell-7.1.3-win-x64.msi' -OutFile 'PowerShell-7.1.3-win-x64.msi'
    Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i PowerShell-7.1.3-win-x64.msi /quiet /norestart' -Wait
    Remove-Item 'PowerShell-7.1.3-win-x64.msi'
    Write-Host 'PowerShell 7.x has been installed.'
}


#Step 3: Create zLibrary and Create Shortcuts



