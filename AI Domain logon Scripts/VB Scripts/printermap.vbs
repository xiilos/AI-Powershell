	'VBScript written by Jay Becker, Advantage International'

'Option Explicit
Dim strUserName, objNetwork, strUNCPrinter, strUNCPrinter2, strUNCPrinter3
Dim strUNCPrinter4, strUNCPrinter5, strUNCPrinter6, strUNCPrinter7, strUNCPrinter8
 

On Error Resume Next






Set objNetwork = CreateObject("WScript.Network") 

strUserName = objNetwork.UserName 



  
'Removed Code - Printer removal function'
'First, remove all old network printers'
'Set WshNetwork2 = WScript.CreateObject("WScript.Network")
'Set Printers = WshNetwork2.EnumPrinterConnections

'For i = 0 to Printers.Count - 1 Step 2
'    If Left(ucase(Printers.Item(i+1)),2) = "\\" Then
'        'WScript.Echo Printers.Item(i+1)
'        WSHNetwork2.RemovePrinterConnection Printers.Item(i+1)
'    End IF
'Next

strUNCPrinter = "\\DFSFILESERVER\HP LaserJet 4000 Series (Receptionist)"
strUNCPrinter2 = "\\DFSFILESERVER\HP LaserJet 4000 Series (Accounting)"
3strUNCPrinter3 = "\\DFSFILESERVER\HP LaserJet P4014/P4015 (Office 1)"
#'strUNCPrinter4 = "\\DC1\HP LaserJet 4000 (Ted)"
#'strUNCPrinter5 = "\\SRV1\HPColorL"
#'strUNCPrinter6 = "\\SRV1\HPLJTC"
#if InStr(strUserName,"tina") Then strUNCPrinter7 = "\\192.168.0.208\HP Color LaserJet 3600"
#if InStr(strUserName,"tina") Then strUNCPrinter8 = "\\WS34\HP LaserJet 4P"


objNetwork.AddWindowsPrinterConnection strUNCPrinter
objNetwork.AddWindowsPrinterConnection strUNCPrinter2
objNetwork.AddWindowsPrinterConnection strUNCPrinter3
#'objNetwork.AddWindowsPrinterConnection strUNCPrinter4
#'objNetwork.AddWindowsPrinterConnection strUNCPrinter5
#'objNetwork.AddWindowsPrinterConnection strUNCPrinter6
#if InStr(strUserName,"tina") Then objNetwork.AddWindowsPrinterConnection strUNCPrinter7
#if InStr(strUserName,"tina") Then objNetwork.AddWindowsPrinterConnection strUNCPrinter8


WScript.Quit



