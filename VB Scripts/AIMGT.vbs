	'VBScript written by Jay Becker, Advantage International'

Option Explicit
Dim objNetwork, strRemotePath1, strRemotePath2, strRemotePath3, strRemotePath4, strRemotePath5, strRemotePath6
Dim strDriveLetter1, strDriveLetter2, strDriveLetter3, strDriveLetter4, strDriveLetter5, strDriveLetter6, strDriveLetter7, strDriveLetter8
Dim strRemotePath7, strRemotePath8

strDriveLetter1 = "F:" 
strDriveLetter2 = "L:" 
strDriveLetter3 = "P:" 
strDriveLetter4 = "K:" 
strDriveLetter5 = "X:"  
strDriveLetter6 = "Y:"
strDriveLetter7 = "Q:"
strDriveLetter8 = "O:"
strRemotePath1 = "\\diditbetter.com\dfs\FileServer" 
strRemotePath2 = "\\diditbetter.com\dfs\Legal" 
strRemotePath3 = "\\diditbetter.com\dfs\HR" 
strRemotePath4 = "\\diditbetter.com\dfs\Marketing" 
strRemotePath5 = "\\diditbetter.com\dfs\Accounting" 
strRemotePath6 = "\\diditbetter.com\dfs"
strRemotePath7 = "\\diditbetter.com\dfs\Quickbooks"
strRemotePath8 = "\\diditbetter.com\dfs\Order Fulfillment"

Set objNetwork = CreateObject("WScript.Network") 


objNetwork.MapNetworkDrive strDriveLetter1, strRemotePath1
objNetwork.MapNetworkDrive strDriveLetter2, strRemotePath2
objNetwork.MapNetworkDrive strDriveLetter3, strRemotePath3
objNetwork.MapNetworkDrive strDriveLetter4, strRemotePath4
objNetwork.MapNetworkDrive strDriveLetter5, strRemotePath5
objNetwork.MapNetworkDrive strDriveLetter6, strRemotePath6
objNetwork.MapNetworkDrive strDriveLetter7, strRemotePath7
objNetwork.MapNetworkDrive strDriveLetter8, strRemotePath8

Include("\\diditbetter\netlogon\PrinterMap.vbs")

Sub Include(sInstFile)
	Dim oFSO, f, s
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set f = oFSO.OpenTextFile(sInstFile)
	s = f.ReadAll
	f.Close	
	ExecuteGlobal s
End Sub



Wscript.Quit


