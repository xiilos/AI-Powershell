if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

(Invoke-RestMethod -Uri https://s3.amazonaws.com/dl.diditbetter.com).OuterXml






$version = (Invoke-RestMethod -Uri https://s3.amazonaws.com/dl.diditbetter.com).OuterXml | Where-Object {$_.Name -like "a2e-standard_upgrade*"} | Select-object -ExpandProperty





| Where-Object {$_.Name -eq 'DefaultEnvironmentId'} | Select-object -ExpandProperty 'Value' | Select-object -ExpandProperty 'Value'

$string = $web.DownloadString("https://s3.amazonaws.com/dl.diditbetter.com")
if ($string -like "a2e-standard_upgrade*") {
    Write-Host "A2E Standard Exists"
} else {
    Write-Host "A2E Standard not here"
}




version_mojoportal.py move release-candidate upgrade %[upgrade_id] %[upgrade_name] %[version] %_date "%@filesize[%[upgrade_name], M] MB"

https://s3.amazonaws.com/dl.diditbetter.com/a2e-standard_upgrade.19.7.185.229.exe
https://s3.amazonaws.com/dl.diditbetter.com/a2e-standard_upgrade.19.7.185.229.exe




# import module
Import-Module AWSPowerShell.NetCore

# Set variable
# To get a list of bucket names use 
Get-S3Bucket -Region us-east-1 # specify the region where your bucket is created

$bucket = "dl.diditbetter.com"

# Get the key of the object you want to query
# For this example I am using the first object returned
$obj = (Get-S3Object -BucketName $bucket)[0].Key



$url = "https://s3.amazonaws.com/dl.diditbetter.com"
Invoke-RestMethod -Method Post -Uri $url -Body $body -ContentType 'application/xml'

$entries = Invoke-RestMethod -Uri "https://s3.amazonaws.com/dl.diditbetter.com"
$entries.title





$sourceBucket = 'dl.diditbetter.com'
$profile = 'DidItBetter'
$Folder = 'c:\'

$items = Get-S3Object -BucketName $sourceBucket -ProfileName $profile -Region 'us-east-1'
Write-Host "$($items.Length) objects to copy"
$index = 1
$items | % {
    Write-Host "$index/$($items.Length): $($_.Key)"
    $fileName = $Folder + ".\$($_.Key.Replace('/','\'))"
    Write-Host "$fileName"
    Read-S3Object -BucketName $sourceBucket -Key $_.Key -File $fileName -ProfileName $profile -Region 'us-east-1' > $null
    $index += 1
}




https://support.diditbetter.com/downloads.aspx




$WebPath = "https://s3.amazonaws.com/dl.diditbetter.com"
$request = [System.Net.WebRequest]::Create( $WebPath ) 
$headers = $request.GetResponse().Headers 
# Content-disposition includes a name-value pair for filename:
$cd = $headers.GetValues("Content-Disposition")
$cd_array = $cd.split(";")
foreach ($item in $cd_array) { 
  if ($item.StartsWith("a2e-enterprise_upgrade")) {
      # Get string after equal sign
      $filename = $item.Substring($item.IndexOf("=")+1)
      # Remove quotation marks, if any
      $filename = $filename.Replace('"','')
  }
}








$URL = "https://s3.amazonaws.com/dl.diditbetter.com/a2e-enterprise_upgrade.24.3.3398.2377.exe"
$Output = "c:\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"











$WebResponse = Invoke-WebRequest "http://support.diditbetter.com/downloads.aspx"
$WebResponse.content

$Response = $WebResponse.content
get-content $Response | select-string 'a2e-enterprise_upgrade'

$Response = $WebResponse.content

$string = $response
if ($string -match '(a2e-enterprise_upgrade)'){ $A2EVersion = $Matches.Value}


Get-WmiObject -List | Where-Object {$_.name  -Like "*a2e-enterprise_upgrade"}



$uri = Invoke-WebRequest -Uri "https://s3.amazonaws.com/dl.diditbetter.com/" -ContentType "application/xml" -ErrorAction:Stop -TimeoutSec 60
$bn = ([xml]$uri.Content).key.lastmodified.ETag


<Contents>
<Key>a2e-enterprise_upgrade.24.3.3398.2377.exe</Key>
<LastModified>2022-03-10T17:12:49.000Z</LastModified>
<ETag>"1c2641b93a39f46c8789b4fd5f95374b"</ETag>
<Size>44214670</Size>
<StorageClass>REDUCED_REDUNDANCY</StorageClass>
</Contents>



$r = Invoke-WebRequest -Uri "$uriP/api/1.0/appliance-management/global/info" -Body $body -Method:Get -Headers $head -ContentType "application/xml" -ErrorAction:Stop -TimeoutSec 60
$bn = ($r.Content.globalInfo.versionInfo.buildNumber)














$url = 'https://s3.amazonaws.com/dl.diditbetter.com/' | Where-Object { $_.DisplayName -like "a2e-enterprise_upgrade*" } | Select-Object
$targetfolder = "c:\zlibrary"
Start-BitsTransfer -Source $url -Destination $targetfolder -Asynchronous -Priority Low









#Downloading Add2Exchange

$URL = "http://dl.diditbetter.com" 
$result = (((Invoke-WebRequest –Uri $url).Links | Where-Object {$_.href -like “a2e_enterprise_upgrade*”} ) | Select-Object href).href






$R = Invoke-WebRequest -URI "http://dl.diditbetter.com"
$R.AllElements | Where-Object {$_.name -like "* a2e_enterprise_upgrade"} | Select-Object Name



(Invoke-WebRequest -Uri "https://s3.amazonaws.com/dl.diditbetter.com" -UseBasicParsing).Links.Href


$URL = "http://dl.diditbetter.com" | Where-Object { $_.name -like "* a2e_enterprise_upgrade*"} | Select-Object
$Output = "c:\zlibrary\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"




$Response = Invoke-WebRequest -URI https://s3.amazonaws.com/dl.diditbetter.com/ | Where-Object { $_.name -like "* a2e_enterprise_upgrade*"} -outfile "C:\zlibrary\a2e-enterprise_upgrade.exe"
$Response.InputFields | Where-Object {
    $_.name -like "* a2e_enterprise_upgrade*"
} | Select-Object Name, Value






Invoke-Webrequest "https://s3.amazonaws.com/dl.diditbetter.com/" | Where-Object {$_.Name -Match '2299.exe._$'} | outfile "C:\zlibrary\a2e-enterprise_upgrade.exe"






Write-Host "Downloading Add2Exchange"
Write-Host "Please Wait......"

$ProgressPreference = 'SilentlyContinue'
Invoke-Webrequest "http://dl.diditbetter.com/a2e-enterprise.22.5.3089.1863.exe" -outfile "C:\zlibrary\a2e-enterprise_upgrade.exe"




Invoke-Webrequest "http://dl.diditbetter.com" | Where-Object {($_.Name -like "Add2Exchange-Enterprise*")} , "C:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"





<#

$URL = "dl.diditbetter.com\a2e-enterprise.22.5.3089.1863.exe"
$Output = "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe"
$Start_Time = Get-Date

(New-Object System.Net.WebClient).DownloadFile($URL, $Output)

Write-Output "Time taken: $((Get-Date).Subtract($Start_Time).Seconds) second(s)"

Write-Host "Finished Downloading"

#Unpacking Add2Exchange

Write-Host "Unpacking Add2exchange"
Write-Host "please Wait....."
Push-Location "c:\zlibrary\Add2Exchange Upgrades"
Start-Process "c:\zlibrary\Add2Exchange Upgrades\a2e-enterprise_upgrade.exe" -wait
Start-Sleep -Seconds 2
Write-Host "Done"

#Installing Add2Exchange
Do {
    Write-Host "Installing Add2Exchange"
    $Location = Get-ChildItem -Path $root | Where-Object { $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    Push-Location $Location
    Start-Process msiexec.exe -Wait -ArgumentList '/I "Add2Exchange_Upgrade.msi" /quiet' -ErrorAction Inquire -ErrorVariable InstallError;
    Write-Host "Finished...Upgrade Complete"

    If ($InstallError) { 
        Write-Warning -Message "Something Went Wrong with the Install!"
        Write-Host "Trying The Install Again in 2 Seconds"
        Start-Sleep -S 2
    }
} Until (-not($InstallError))



#Setting the Service Account Password
$Password = Get-Content "C:\Program Files (x86)\DidItBetterSoftware\Add2Exchange Creds\Local_Account_Pass.txt" | convertto-securestring

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password)
$SAP = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)

$SVC = Get-WmiObject win32_service -Filter "Name='Add2Exchange Service'"
$SVC.StopService();
$Result = $SVC.Change($Null, $Null, $Null, $Null, $Null, $Null, $Null, "$SAP")
If ($Result.ReturnValue -eq '0') { Write-Host "Add2Exchange Service Password Has Been Succsefully Updated" -ForegroundColor Green } Else { Write-Host "Error: $Result" }



#Setting the Add2Exchange Service to Delayed Start
Write-Host "Setting up Add2Exchange Service to Delayed Start"
sc.exe config "Add2Exchange Service" start= delayed-auto
Write-Host "Done"
#>
Pause

Write-Host "ttyl"
Get-PSSession | Remove-PSSession
Exit

# End Scripting