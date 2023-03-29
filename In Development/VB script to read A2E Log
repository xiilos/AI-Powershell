On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery _
("Select * from Win32_NTLogEvent Where LogFile = 'Add2Exchange' and EventCode = '1234' or EventCode = '5678'")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.OpenTextFile("C:\Add2ExchangeLog.txt", ForWriting, True)

For Each objEvent in colLoggedEvents
    objOutputFile.WriteLine "Event Date: " & objEvent.TimeGenerated
    objOutputFile.WriteLine "Event ID: " & objEvent.EventCode
    objOutputFile.WriteLine "Message: " & objEvent.Message
    objOutputFile.WriteLine
Next

objOutputFile.Close
