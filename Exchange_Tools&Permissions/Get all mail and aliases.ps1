if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force


# Script #

get-mailbox | ForEach-Object{ 

  $host.UI.Write("Blue", $host.UI.RawUI.BackGroundColor, "`nUser Name: " + $_.DisplayName+"`n")
 
  for ($i=0;$i -lt $_.EmailAddresses.Count; $i++)
  {
     $address = $_.EmailAddresses[$i]
     
     $host.UI.Write("Blue", $host.UI.RawUI.BackGroundColor, $address.AddressString.ToString()+"`t")
  
     if ($address.IsPrimaryAddress)
     { 
       $host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Primary Email Address`n")
     }
    else
    {
       $host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "Alias`n")
     }
  }
 } | out-file "C:\Email Addresses.txt"

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting