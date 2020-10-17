	'VBScript written by Jay Becker, Advantage International'

Option Explicit
Dim objNetwork, strRemotePath1, strRemotePath2, strRemotePath3
Dim strDriveLetter1, strDriveLetter2, strDriveLetter3 

strDriveLetter1 = "F:" 
'strDriveLetter2 = "Y:" 
strDriveLetter3 = "O:"
strRemotePath1 = "\\diditbetter.com\DFS\FileServer" 
'strRemotePath2 = "\\diditbetter.com\DFS" 
strRemotePath3 = "\\diditbetter.com\DFS\Order Fulfillment" 

On Error Resume Next

Set objNetwork = CreateObject("WScript.Network") 

' Section which maps two drives, M: and P:
objNetwork.MapNetworkDrive strDriveLetter1, strRemotePath1
objNetwork.MapNetworkDrive strDriveLetter2, strRemotePath2
objNetwork.MapNetworkDrive strDriveLetter3, strRemotePath3

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


