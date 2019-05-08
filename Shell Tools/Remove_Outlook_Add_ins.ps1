if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
  
#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass
  
#Menu
$Title1 = 'Disabling Outlook Social Connector'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "What Version of Outlook?"
""
Write-Host "Press '1' for Outlook 2010"
Write-Host "Press '2' for Outlook 2013"
Write-Host "Press '3' for Outlook 2016, O365, or Outlook 2019"
Write-Host "Press 'Q' to Quit." -ForegroundColor Red


#OutLook Version
 
$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    # Remove Outlook Social Connector for 2010 x86-x64
    '1' { 
        Clear-Host 
        'You chose Outlook 2010'
        $TestPath = "C:\Program Files (x86)\Microsoft Office\Office14"
  
        If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
            Write-Host "32bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

        }
        Else {
            Write-Host "64bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

        }
    }


    # Remove Outlook Social Connector for 2013 x86-x64
    '2' {
        Clear-Host 
        'You chose Outlook 2013'
        $TestPath = "C:\Program Files (x86)\Microsoft Office\Office15"
  
        If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
            Write-Host "32bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

        }
        Else {
            Write-Host "64bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null

        }
        
    }
    # Remove Outlook Social Connector for 2016, 365, & 2019 x86-x64
    '3' {
        Clear-Host 
        'You chose Outlook 2016, O365, or Outlook 2019'
        $TestPath = "C:\Program Files (x86)\Microsoft Office\root\Office16"
  
        If ( $(Try { Test-Path $TestPath.trim() } Catch { $false }) ) {
            Write-Host "32bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null
    
        }
        Else {
            Write-Host "64bit Outlook"
            Write-Host "Setting Load Behavior to 0"
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins\OscAddin.Connect' -Name LoadBehavior -Value 0 | Out-Null
    
        }

    }

    # Option Q:-Quit
    'q' { 
        Write-Host "Quitting"
        Get-PSSession | Remove-PSSession
        Exit 
    }
}

Write-Host "Done"
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please Restart Outlook", 0, "Done", 0x1)
Start-Sleep 1
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit
  
# End Scripting