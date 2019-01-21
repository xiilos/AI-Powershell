if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  # Relaunch as an elevated process:
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

#Execution Policy

#Set-ExecutionPolicy -ExecutionPolicy Unrestricted


# Script #

#Folder to find:
$Foldername = Write-Host "Name of Folder to Remove"
      
foreach ($Mailbox in Get-Mailbox -resultsize unlimited)
{

$Emailaddress = $Mailbox.primarysmtpaddress

#Connect to EWS
Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$Service.AutodiscoverUrl($EmailAddress,{$true})
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
$service.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)

#Bind to the Mesasge Root
$RootFolderName = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$EmailAddress)
$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RootFolderName)
$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep 
$Folders = $RootFolder.FindFolders($FolderView)

$FolderFound = $Folders | Where-Object{$_.displayname -Like $Foldername}

#You can choose from a few delete types, just choose one:
$FolderFound.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
#$FolderFound.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
#$FolderFound.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)

}


Write-Host "Done"
Write-Host "Quitting"
Get-PSSession | Remove-PSSession
Exit

# End Scripting