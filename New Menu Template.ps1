if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #
$Title1 = 'Main Menu Title'

Clear-Host 
Write-Host "================ $Title1 ================"
""
Write-Host "How Are We Logging In?"
""
Write-Host "Press '1' for Office 365"
Write-Host "Press '2' for Exchange 2010" 
Write-Host "Press '3' for Exchange 2013-2016" 
Write-Host "Press 'Q' to Quit." -ForegroundColor Red


#Login Method
 
$input1 = Read-Host "Please Make A Selection" 
switch ($input1) { 

    #Office 365--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '1' { 
        Clear-Host 
        'You chose Office 365'
    }
  }
            $Title2 = 'Menu 2' 
            ""
            Clear-Host 
            Write-Host "================ $Title2 ================" 
            ""     
            Write-Host "Press '1' for Adding Permissions to All Users" 
            Write-Host "Press '2' for Removing Permissions from All Users" 
            Write-Host "Press '3' for Removing and then Adding Permissions to All Users"
            Write-Host "Press '4' for Adding Permissions to a Distribution List"
            Write-Host "Press '5' for Removing Permissions from a Distribution List"
            Write-Host "Press '6' for Adding Permissions to a Single User"
            Write-Host "Press '7' for Removing Permissions from a Single User"
            Write-Host "Press '8' for Adding Permissions to Public Folders"
            Write-Host "Press '9' for Removing Permissions From Public Folders" 
            Write-Host "Press 'Q' to Quit" -ForegroundColor Red
    
            
            $input2 = Read-Host "Please Make A Selection" 
            switch ($input2) {

                # Option 1: Office 365-Adding Add2Exchange Permissions
                '1' { 
                    Clear-Host 
                    'You chose to Add Permissions to All Users'
                }
              }


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting