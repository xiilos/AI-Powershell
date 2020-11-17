'==========================================================================
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
' NAME: Create_new_local_admin_account.vbs
' AUTHOR: David Taylor, E.W. Scripps Company 
' DATE  : 9/9/2013
' COMMENT: This script creates a new local admin account, sets the password
' to never expire, and disables the built-in Administrator account.
'==========================================================================
'Add "Admin" account to Users group
' Change "." to remote computer name on line 11 if needed.
strComputer = "."
Set colAccounts = GetObject("WinNT://" & strComputer & "")
' Change "New_local_Administrator_account_name" in line 14 to the account name you wish to use.
Set objUser = colAccounts.Create("user", "New_local_Administrator_account_name")
' Change "NewPassword" in line 16, to the new password you wish to use. 
objUser.SetPassword "NewPassword"
objUser.SetInfo
'Pause script for 10 seconds before continuing
'WScript.sleep 10000
'Add "New_local_Administrator_account_name" account to Administrators group
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
' Change New_local_Administrator_account_name in line 23 to the account name in line 14.
Set objUser = GetObject("WinNT://" & strComputer & "/New_local_Administrator_account_name,user")
objGroup.Add(objUser.ADsPath)
'Set "New_local_Administrator_account_name" account's Password to never expire
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
' Change "DomainName" in line 28 to reflect your domain.
strDomainOrWorkgroup = "DomainName"
' Change "New_local_Administrator_account_name" in line 30 to the account name defined in line 13.
strUser = "New_local_Administrator_account_name"
Set objUser = GetObject("WinNT://" & strDomainOrWorkgroup & "/" & _
    strComputer & "/" & strUser & ",User")
objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo
'Disable the "Administrator" account
Set objUser = GetObject("WinNT://" & strComputer & "/Administrator")
objUser.AccountDisabled = True
objUser.SetInfo