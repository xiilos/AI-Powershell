	'VBScript written by Jay Becker, Advantage International'

Option Explicit
Dim objNetwork
Set objNetwork = CreateObject("WScript.Network") 


On Error Resume Next 


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run "wscript.exe \\collection\netlogon\PrinterMap.vbs param1", , True
WSHShell.Run "wscript.exe \\collection\netlogon\DriveMaps.vbs", , True


'Set Default printer based on location'
'objNetwork.SetDefaultPrinter "\\ws23\HPLaser"





Wscript.Quit

