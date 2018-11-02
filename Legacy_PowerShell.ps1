if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Check if .Net 4.5 or above is installed

$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389))
{
    Write-Host "We need to download .Net 4.5.2"
    Write-Host "Downloading"
    Write-Host "You must reboot after install and run this again"
    $url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    $output = "c:\zlibrary\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Write-Host "Download Complete"
    
    Invoke-item -Path "c:\zlibrary\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
}


#Check Operating Sysetm

$BuildVersion = [System.Environment]::OSVersion.Version


#OS is 10+
if($BuildVersion.Major -like '10')
    {
        Write-Host "WMF 5.1 is not supported for Windows 10 and above"
        
    }

#OS is 7
if($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1')
    {
        
Write-Host "Downloading WMF 5.1 for 7+"
Write-Host "You must reboot after install"
$url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
$output = "c:\zlibrary\Win7AndW2K8R2-KB3191566-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Write-Host "Download Complete"
    
    Invoke-Item -Path 'c:\zlibrary\Win7AndW2K8R2-KB3191566-x64.msu'
    }

#OS is 8
elseif($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3')
    {
        
Write-Host "Downloading WMF 5.1 for 8+"
Write-Host "You must reboot after install"
$url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
$output = "c:\zlibrary\Win8.1AndW2K12R2-KB3191564-x64.msu"
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Write-Host "Download Complete"
    
    Invoke-Item -Path 'c:\zlibrary\Win8.1AndW2K12R2-KB3191564-x64.msu'
    }

Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell

$wshell.Popup("Please Reboot after Installing",0,"Done",0x1)
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting