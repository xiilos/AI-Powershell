if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append

#Create zLibrary\A2E Diags Directory

Write-Host "Creating Landing Zone"
$TestPath = "C:\zlibrary\A2E Diags"
if ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {

    Write-Host "A2E Diags Directory Exists...Resuming"
}
Else {
    New-Item -ItemType directory -Path "C:\zlibrary\A2E Diags"
}

#Downloading A2E Diags

Write-Host "Downloading A2E Diags"
Write-Host "Please Wait......"

$URL = "ftp://ftp.diditbetter.com/A2EDiags/A2EDiags-2.3.exe"
$Output = "C:\zlibrary\A2E Diags\A2EDiags-2.3.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking A2E Diags

Write-Host "Unpacking A2E Diags"
Write-Host "please Wait....."
Push-Location "C:\zlibrary\A2E Diags"
Start-Process -FilePath "./A2EDiags-2.3.exe" -wait -ErrorAction Stop
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\A2EDiags.lnk")
$Shortcut.TargetPath = "C:\zlibrary\A2E Diags\A2EDiags-2.3\A2EDiags.cmd"
$Shortcut.WorkingDirectory = "C:\zlibrary\A2E Diags\A2EDiags-2.3"
$Shortcut.Save()
Write-Host "Finished..."

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting