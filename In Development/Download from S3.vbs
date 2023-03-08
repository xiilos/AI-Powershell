' Set up variables for the S3 bucket and file name
Dim s3Bucket, partialFileName
s3Bucket = "downloads.diditbetter.com"
partialFileName = "a2e-enterprise_upgrade"

' Set the path where the file will be saved
savePath = "C:\zlibrary\Add2Exchange Upgrades\"

' Create a new XMLHTTP object
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")

' Construct the URL for the S3 file using the AWS S3 REST API
Dim s3Url
s3Url = "https://s3.amazonaws.com/" & s3Bucket & "/"

' Send a GET request to the S3 URL
xmlhttp.Open "GET", s3Url, False
xmlhttp.send

' Parse the response text into an XML object
Dim responseXml
Set responseXml = CreateObject("MSXML2.DOMDocument")
responseXml.LoadXML(xmlhttp.responseText)

' Find the file with the partial name in the XML response
Dim matchingNode
For Each node In responseXml.getElementsByTagName("Key")
    If InStr(node.Text, partialFileName) > 0 Then
        Set matchingNode = node
        Exit For
    End If
Next

' If a matching file was found, download it using another HTTP request
If Not matchingNode Is Nothing Then
    Dim fileUrl
    fileUrl = s3Url & matchingNode.Text

    ' Send another GET request to the file URL to download it
    xmlhttp.Open "GET", fileUrl, False
    xmlhttp.send

    ' Save the downloaded file to disk
    Dim fileStream
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' binary
    fileStream.Open
    fileStream.Write xmlhttp.responseBody
    fileStream.SaveToFile savepath & matchingNode.Text, 2 ' overwrite
    fileStream.Close
End If
