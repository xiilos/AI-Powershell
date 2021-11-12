strFolderPath = "Public Folders - zAdd2Exchange@devb.Local\All Public Folders\Contacts\Firm Contacts"
aryFolders = Split(strFolderPath, "\")
Set olApp = CreateObject("Outlook.Application")
Set olNS = olApp.GetNamespace("MAPI")
Set olFolder = olNS.Folders.Item(aryFolders(0))
If Not olFolder Is Nothing Then
                For lngLoop = 1 To UBound(aryFolders)
                                Set olFolders = olFolder.Folders
                                Set olFolder = Nothing
                                Set olFolder = olFolders.Item(aryFolders(lngLoop))
                                If olFolder Is Nothing Then Exit For
                Next
End If 'Folder Set

If Not olFolder Is Nothing Then
                If olFolder.DefaultItemType = 2 And StrComp(strFolderPath, Mid(olFolder.FolderPath, 3), vbTextCompare) = 0 Then 'olContactItem = 2
                                olFolder.ShowAsOutlookAB = True
                                Result = MsgBox(olFolder.FolderPath & " Show As Outlook Address Book Set", 0, "Completed")
                Else
                                Result = MsgBox(olFolder.FolderPath & " Show As Outlook Address Book Failed", 16, "Failed")
                End If 'Folder Contact And Folder Path Match
Else
                Result = MsgBox(strFolderPath  & " Folder Set Failed", 16, "Failed")
End If 'Folder Set