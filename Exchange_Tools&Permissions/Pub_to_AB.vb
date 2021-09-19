Sub OutlookFolderShowAsOutlookAB()
    Dim olFolder As Outlook.MAPIFolder
    Dim strFolderPath As String
    
    On Error Resume Next
    strFolderPath = "Public Folders - zAdd2Exchange@DevA.Local\All Public Folders\A2E GAL Cache"
    Set olFolder = OutlookGetFolderFromPath(strFolderPath)
    If Not olFolder Is Nothing Then
        If olFolder.DefaultItemType = olContactItem And StrComp(strFolderPath, Mid(olFolder.FolderPath, 3), vbTextCompare) = 0 Then
            olFolder.ShowAsOutlookAB = True
            Debug.Print olFolder.FolderPath & " Shown As Address Book Is " & CStr(olFolder.ShowAsOutlookAB)
        End If 'Folder Contact And Folder Path Match
    End If 'Folder Set
    Err.Clear
End Sub
Function OutlookGetFolderFromPath(strFolderPath As String) As Outlook.MAPIFolder
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olFolders As Outlook.Folders
    Dim olFolder As Outlook.Folder
    Dim aryFolders() As String
    Dim lngLoop As Long
    
    On Error Resume Next
    aryFolders() = Split(strFolderPath, "\")
    Set olApp = CreateObject("Outlook.Application")
    Set olNS = olApp.GetNamespace("MAPI")
    Set olFolder = olNS.Folders.Item(aryFolders(0))
    If Not olFolder Is Nothing Then
        For lngLoop = 1 To UBound(aryFolders)
            Set olFolders = olFolder.Folders
            Set olFolder = Nothing
            Set olFolder = olFolders.Item(aryFolders(lngLoop))
            If olFolder Is Nothing Then Exit For
        Next lngLoop
        Set OutlookGetFolderFromPath = olFolder
        Set olFolder = Nothing
        Set olFolders = Nothing
        Set olNS = Nothing
        Set olApp = Nothing
    End If 'Folder Set
    Err.Clear
End Function
