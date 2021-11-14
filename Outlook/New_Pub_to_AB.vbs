Option Explicit

Dim objOL, objNS, objFolder
Set objOL = CreateObject("Outlook.application")
Set objNS = objOL.GetNamespace("MAPI")


Set objFolder = objNS.GetDefaultFolder(18).Folders("Contacts").Folders("Firm Contacts")
    objFolder.ShowAsOutlookAB = True



    '18 is olPublicFoldersAllPublicFolders representing all public folders