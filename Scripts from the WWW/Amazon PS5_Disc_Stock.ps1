
# Script #

#function Watch-Amazon
{
 param
  (
 #   [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)][string]$ASIN,
    [parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)][string]$Stock = "in stock"
#    [parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $true)][string]$Interval = 3600
  )
$ASIN = "B08FC5L3RG"


$query="http://www.amazon.com/dp/" + $ASIN
  $WebClient = New-Object System.Net.WebClient

  While ($true)
  {
    $Amazon = $WebClient.DownloadString($query)
    if ($Amazon -match $Stock)
      { 
        #Send-MailMessage -From "AmazonWatch@domain.com" -To "user@domain.com" -Subject "Your product is now in stock at Amazon!" -Body "Get it now: $query"
        Write-Host "It's in stock!" -ForegroundColor green
      }
    else
      { 
        Write-Host  "Not in stock yet... please wait!" -ForegroundColor red
      }
    Start-sleep $interval
  }
}

#Watch-Amazon



#Playstation 5 Disc Version
# ASIN = B08FC5L3RG



# End Scripting