if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Variables
$Domain = (Get-ADDomain).DNSRoot
$Partition = Get-ADDomainController -Filter * -Server $Domain | Select-Object -ExpandProperty DefaultPartition


<#Checking PowerShell Version

$wshell = New-Object -ComObject Wscript.Shell
    
$answer = $wshell.Popup("Caution... We must first check to ensure you are on the latest Powershell Version. This may require a reboot. Click OK to Continue, or Cancel to Quit ", 0, "WARNING!!", 0x1)
if ($answer -eq 2) {
    Write-Host "ttyl"
    Get-PSSession | Remove-PSSession
    Exit
}

$release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Release -ErrorAction SilentlyContinue -ErrorVariable evRelease).release
$installed = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name Install -ErrorAction SilentlyContinue -ErrorVariable evInstalled).install

if (($installed -ne 1) -or ($release -lt 378389)) {
    Write-Host "We need to download .Net 4.5.2"
    Write-Host "Downloading"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    $output = "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)    
    Start-Process -FilePath "C:\PowerShell\NDP452-KB2901907-x86-x64-AllOS-ENU.exe" -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing and run this again", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}


#Check Operating Sysetm

$BuildVersion = [System.Environment]::OSVersion.Version


#OS is 10+
if ($BuildVersion.Major -like '10') {
    Write-Host "WMF 5.1 is not supported for Windows 10 and above"
        
}

#OS is 7
if ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '1') {
        
    Write-Host "Downloading WMF 5.1 for 7+"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/Win7AndW2K8R2-KB3191566-x64.msu"
    $output = "C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Start-Process -FilePath 'C:\PowerShell\Win7AndW2K8R2-KB3191566-x64.msu' -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

#OS is 8
elseif ($BuildVersion.Major -eq '6' -and $BuildVersion.Minor -le '3') {
        
    Write-Host "Downloading WMF 5.1 for 8+"
    $Directory = "C:\PowerShell"

    if ( -Not (Test-Path $Directory.trim() )) {
        New-Item -ItemType directory -Path C:\PowerShell
    }
    Write-Host "Downloading the Latest! Please Wait....."
    $url = "ftp://ftp.diditbetter.com/PowerShell/Win8.1AndW2K12R2-KB3191564-x64.msu"
    $output = "C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu"
    (New-Object System.Net.WebClient).DownloadFile($url, $output)
    Start-Process -FilePath 'C:\PowerShell\Win8.1AndW2K12R2-KB3191564-x64.msu' -wait
    Write-Host "Download Complete"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Please Reboot after Installing", 0, "Done", 0x1)
    Write-Host "Quitting"
    Get-PSSession | Remove-PSSession
    Exit
}

Write-Host "You Are on the latest version of PowerShell"


#>
#---------------------------------------------------------------------------------------------------------------------------------------


# Script #

Write-Host "What OS Version is the DC on?"
""
Write-Host "Press '1' for Server 2008 R2"
Write-Host "Press '2' for 2012 R2"
Write-Host "Press '3' for 2016+"
Write-Host "Press 'Q' to Quit." -ForegroundColor Red


#DC OS Version
 
$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    #DC Options--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '1' { 
        Clear-Host 
        'You chose Server 2008R2'
        #Create the GPO
        Import-GPO -BackupId 15D2F64C-C68B-4DE7-8647-86F0D6B149CE -TargetName "Disable Outlook Social Connector" -path "./Server2008R2" -CreateIfNeeded -Domain $Domain
        $GPO = (get-gpo -name "Disable Outlook Social Connector" | Select-Object -ExpandProperty ID | Format-Table -hidetableheaders | out-String).Trim()
        Write-Host "Copying Files Needed..."
        Copy-Item ".\OSC_Disable.bat" -Destination "C:\Windows\SYSVOL\sysvol\$Domain\Policies\{$GPO}\Machine\Scripts\Startup"
        Set-GPPermissions -Name "Disable Outlook Social Connector" -TargetName "Domain Computers" -TargetType Group -PermissionLevel GpoRead -ErrorAction SilentlyContinue
        Set-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition" -LinkEnabled Yes -ErrorAction SilentlyContinue -ErrorVariable LinkFault
        If ($LinkFault) { 
            New-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition"
        }
        Write-Host "Done"
    }

    '2' { 
        Clear-Host 
        'You chose Server 2012R2'
        #Create the GPO
        Import-GPO -BackupId F8903325-19D1-461D-B306-60783653DEAF -TargetName "Disable Outlook Social Connector" -path "./2012R2" -CreateIfNeeded -Domain $Domain
        $GPO = (get-gpo -name "Disable Outlook Social Connector" | Select-Object -ExpandProperty ID | Format-Table -hidetableheaders | out-String).Trim()
        Write-Host "Copying Files Needed..."
        Copy-Item ".\OSC_Disable.bat" -Destination "C:\Windows\SYSVOL\sysvol\$Domain\Policies\{$GPO}\Machine\Scripts\Startup"
        Set-GPPermissions -Name "Disable Outlook Social Connector" -TargetName "Domain Computers" -TargetType Group -PermissionLevel GpoRead -ErrorAction SilentlyContinue
        Set-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition" -LinkEnabled Yes -ErrorAction SilentlyContinue -ErrorVariable LinkFault
        If ($LinkFault) { 
            New-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition"
        }
        Write-Host "Done"
    }

    '3' { 
        Clear-Host 
        'You chose Server 2016+'
        #Create the GPO
        Import-GPO -BackupId 1CBF8955-FA37-4588-928E-3B76F370422F -TargetName "Disable Outlook Social Connector" -path "./Server2016+" -CreateIfNeeded -Domain $Domain
        $GPO = (get-gpo -name "Disable Outlook Social Connector" | Select-Object -ExpandProperty ID | Format-Table -hidetableheaders | out-String).Trim()
        Write-Host "Copying Files Needed..."
        Copy-Item ".\OSC_Disable.bat" -Destination "C:\Windows\SYSVOL\sysvol\$Domain\Policies\{$GPO}\Machine\Scripts\Startup"
        Set-GPPermissions -Name "Disable Outlook Social Connector" -TargetName "Domain Computers" -TargetType Group -PermissionLevel GpoRead -ErrorAction SilentlyContinue
        Set-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition" -LinkEnabled Yes -ErrorAction SilentlyContinue -ErrorVariable LinkFault
        If ($LinkFault) { 
            New-GPLink -Name "Disable Outlook Social Connector" -Target "$Partition"
        }
        Write-Host "Done"
    }

}


Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting