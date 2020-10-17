sNewPassword = "Work_smarter"

Set oWshNet = CreateObject("WScript.Network")
sComputer = oWshNet.ComputerName
sAdminName = GetAdministratorName

On Error Resume Next
Set oUser = GetObject("WinNT://" & sComputer & "/" & sAdminName & ",user")
oUser.SetPassword sNewPassword
oUser.SetInfo
On Error Goto 0


Function GetAdministratorName()

   Dim sUserSID, oWshNetwork, oUserAccount

   Set oWshNetwork = CreateObject("WScript.Network")
   Set oUserAccounts = GetObject( _
        "winmgmts://" & oWshNetwork.ComputerName & "/root/cimv2") _
        .ExecQuery("Select Name, SID from Win32_UserAccount" _
      & " WHERE Domain = '" & oWshNetwork.ComputerName & "'")

   On Error Resume Next
   For Each oUserAccount In oUserAccounts
     If Left(oUserAccount.SID, 9) = "S-1-5-21-" And _
        Right(oUserAccount.SID, 4) = "-500" Then
       GetAdministratorName = oUserAccount.Name
       Exit For
     End if
   Next
End Function
