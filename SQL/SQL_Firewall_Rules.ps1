if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass

#Logging
Start-Transcript -Path "C:\Program Files (x86)\DidItBetterSoftware\Support\A2E_PowerShell_log.txt" -Append


# Script #
Do {
    $Title1 = 'Firewall Rules for Remote Add2Exchange SQL'

    Clear-Host 
    Write-Host "================ $Title1 ================"
    ""
    Write-Host "Please Pick Were to Apply Firewall Rules"
    ""
    Write-Host "Press '1' for This Local Machine"
    Write-Host "Press '2' for The Remote SQL Server" 
    Write-Host "Press '3' for Another Remote Server" 
    Write-Host "Press 'Q' to Quit." -ForegroundColor Red


    #Login Method
 
    $input1 = Read-Host "Please Make A Selection" 
    switch ($input1) { 

        #This Local Machine--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '1' { 
            Clear-Host 
            'You chose This Local Machine'
            Write-Host "Checking for Previous Firewall Rules"
            Get-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -ErrorAction SilentlyContinue -ErrorVariable NetRule;
            if ($NetRule) {
                Write-Host "Inbound Rule Does not Exist"
                Write-Host "Creating SQL Inboud Firewall Rule"
                New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
                Write-Host "Done"
            } 
        
            Else {
                Write-Host "Add2Exchange Allow SQL Inbound Rule Already Exisits"
            }
        
            $confirmation = Read-Host "Would you like to Set the Inbound Rule for SQL Brower? [Y/N]"
            if ($confirmation -eq 'y') {
                New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL Browser" -Direction Inbound -LocalPort 1434 -Protocol UDP -Action Allow -ErrorAction SilentlyContinue
                Write-Host "Done"
            }

            if ($confirmation -eq 'n') {
                Write-Host "Resuming"
            }
        }
        #Remote SQL Server--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '2' {
            Clear-Host 
            'You chose A Remote SQL Server'
            $SQLServ = Read-Host "What is the name of your SQL Server?"
        
            $confirmation = Read-Host "Would you also like to Set the Inbound Rule for SQL Brower? [Y/N]"
            if ($confirmation -eq 'y') {
                Write-Host "Creating Inbound Rules for Add2Exchange SQL"
                Invoke-Command -ComputerName $SQLServ {
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL Browser" -Direction Inbound -LocalPort 1434 -Protocol UDP -Action Allow -ErrorAction SilentlyContinue
                } -Credential (Get-Credential)
                Write-Host "Done"
            }

            if ($confirmation -eq 'n') {
                Write-Host "Creating Inbound Rules for Add2Exchange SQL"
                Invoke-Command -ComputerName $SQLServ {
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
                } -Credential (Get-Credential)
                Write-Host "Done"
            }

        }

        #Another Remote Server--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '3' {
            Clear-Host 
            'You chose A Remote SQL Server'
            $A1Serv = Read-Host "What is the name of your Server?"
            $confirmation = Read-Host "Would you also like to Set the Inbound Rule for SQL Brower? [Y/N]"
            if ($confirmation -eq 'y') {
                Write-Host "Creating Inbound Rules for Add2Exchange SQL"
                Invoke-Command -ComputerName $A1Serv {
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL Browser" -Direction Inbound -LocalPort 1434 -Protocol UDP -Action Allow -ErrorAction SilentlyContinue
                } -Credential (Get-Credential)
                Write-Host "Done"
            }

            if ($confirmation -eq 'n') {
                Write-Host "Creating Inbound Rules for Add2Exchange SQL"
                Invoke-Command -ComputerName $A1Serv {
                    New-NetFirewallRule -DisplayName "Add2Exchange Allow SQL" -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue
                } -Credential (Get-Credential)
                Write-Host "Done"
            }
        }

        # Option Q: Quit
        'q' { 
            Write-Host "ttyl"
            Get-PSSession | Remove-PSSession
            Exit 
        }

    }
    $repeat = Read-Host 'Do you want to run it again? [Y/N]'
} Until ($repeat -eq 'n')

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting