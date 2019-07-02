if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass


# Script #

#Unique Variables
$FTPServer = "ftp:/ftp.diditbetter.com/"
 {
    $request = [Net.WebRequest]::Create($url)
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
    $response = $request.GetResponse()
    $reader = New-Object IO.StreamReader $response.GetResponseStream() 
    $reader.ReadToEnd()
    $reader.Close()
    $response.Close()
}

#Use the root directory of the FTP Server
$ListFiles = $FTPServer
$files = ($ListFiles -split "`r`n")

Write-Host 'The Files on the FTP Server : '
Write-Host `r
$files 


Pause
Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting