'On Error Resume Next

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ADS_UF_SCRIPT = &H00000001
Const ADS_UF_ACCOUNTDISABLE = &H00000002
Const ADS_UF_HOMEDIR_REQUIRED = &H00000008
Const ADS_UF_LOCKOUT = &H00000010
Const ADS_UF_PASSWD_NOTREQD = &H00000020
Const ADS_UF_PASSWD_CANT_CHANGE = &H00000040
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H00000080
Const ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H00000100
Const ADS_UF_NORMAL_ACCOUNT = &H00000200
Const ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H00000800
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H00001000
Const ADS_UF_SERVER_TRUST_ACCOUNT = &H00002000
Const ADS_UF_DONT_EXPIRE_PASSWD = &H00010000
Const ADS_UF_MNS_LOGON_ACCOUNT = &H00020000
Const ADS_UF_SMARTCARD_REQUIRED = &H00040000
Const ADS_UF_TRUSTED_FOR_DELEGATION = &H00080000
Const ADS_UF_NOT_DELEGATED = &H00100000
Const ADS_UF_USE_DES_KEY_ONLY = &H00200000
Const ADS_UF_DONT_REQUIRE_PREAUTH = &H00400000
Const ADS_UF_PASSWORD_EXPIRED = &H00800000
Const ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = &H01000000
Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_DELETE = 4
Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_UPDATE = 2
Const stTitle = "Creation of KNIGHT Active Directory user accounts"

bInteractive = False
stTradBack = ""
stMiddleIn = ""
stDescription = ""

Set oArgs = Wscript.Arguments
Set oFso = Wscript.CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
stCurrentFolder = oShell.CurrentDirectory
Set oNet = WScript.CreateObject("WScript.Network")
stNBDomain = oNet.UserDomain
Select Case oArgs.Count
	Case 0
		bInteractive = True
	Case 1
		If InStr(1, oArgs(0), "?", vbBinarycompare) Then Call Usage()
	Case 4
		If InStr(1, oArgs(0), "S", vbTextcompare) = 1 Then
			stUserService = GetUserType(oArgs(0))
			stFirstName = ChangeCaseToTitle(oArgs(1))
			stLastName = ChangeCaseToTitle(oArgs(2))
			stDescription = CStr(oArgs(3))
	'		Wscript.Echo "Got data:" & vbCrlf & stUserService & " " & stFirstName & " " & stLastName
		Else
			Call Usage()
		End If
	Case 5
		stUserService = GetUserType(oArgs(0))
		stTradBack = GetUserType(oArgs(1))
		stFirstName = ChangeCaseToTitle(oArgs(2))
		stLastName = ChangeCaseToTitle(oArgs(3))
		stDescription = CStr(oArgs(4))
	'	Wscript.Echo "Got data:" & vbCrlf & stUserService & " " & stTradBack & " " & stFirstName & " " & stLastName
	Case 6
		stUserService = GetUserType(oArgs(0))
		stTradBack = GetUserType(oArgs(1))
		stFirstName = ChangeCaseToTitle(oArgs(2))
		stMiddleIn = ChangeCaseToTitle(oArgs(3))
		stLastName = ChangeCaseToTitle(oArgs(4))
		stDescription = CStr(oArgs(5))
	'	Wscript.Echo "Got data:" & vbCrlf & stUserService & " " & stTradBack & " " & stFirstName & " " & stMiddleIn & " " & stLastName
	Case Else
		Call Usage()
End Select

If bInteractive Then
	stUserService = GetUserType(InputBox("Please, enter account type USER or SERVICE", stTitle, "User"))
	If StrComp(stUserService, "service", vbTextcompare) = 0 Then
		stTemp = InputBox("Please, enter service account Name", stTitle, "Application User")
		If stTemp = "" Then Call Usage()
		arTemp = Split(stTemp, " ", -1, vbBinarycompare)
		If UBound(arTemp) > 0 Then
			stFirstName = ChangeCaseToTitle(arTemp(0))
			stLastName = ChangeCaseToTitle(arTemp(1))
		Else
			stFirstName = ChangeCaseToTitle(stTemp)
			stLastName = ""
		End If
	Else
		stTemp = ""
		stTradBack = GetUserType(InputBox("Please, type category of user account TRADER or BACKOFFICE", stTitle, "Backoffice"))
		stTemp = InputBox("Please, enter user First name, Middle Initial and Last Name", stTitle, "John J Doe")
		If stTemp = "" Then Call Usage()
		arTemp = Split(stTemp, " ", -1, vbBinarycompare)
	'	Wscript.Echo "User name is: " & stTemp & vbTab & "arTemp size is: " & UBound(arTemp)
		Select Case UBound(arTemp)
			Case 1
				stFirstName = ChangeCaseToTitle(arTemp(0))
				stLastName = ChangeCaseToTitle(arTemp(1))
			Case 2
				stFirstName = ChangeCaseToTitle(arTemp(0))
				stMiddleIn = ChangeCaseToTitle(arTemp(1))
				stLastName = ChangeCaseToTitle(arTemp(2))
			Case Else
				Call Usage()
		End Select
	End If
	stDescription = InputBox("Please, enter user account description", stTitle, "Account Description")
	If stDescription = "" Then Call Usage()
End If

Set oRootDSE = GetObject("LDAP://rootDSE")
stRootNC = CStr(oRootDSE.Get("defaultNamingContext"))
stDomain = "LDAP://" & stRootNC
'Wscript.Echo "Default domain naming context: " & stDomain
Call SetAccountProperties(stUserService, stTradBack, stFirstName, stMiddleIn, stLastName, stDescription, bInteractive, stUserNC, stDispName, stUserName, stServer)

Set oUserOU = GetObject(stUserNC)
Set oUser = oUserOU.Create("user", "cn=" & stUserName)
oUser.samAccountName = stUserName
oUser.SetInfo
oUser.description = stDescription
oUser.givenName = stFirstName
oUser.sn = stLastName
If Not (stMiddleIn = "") Then oUser.initials = stMiddleIn
oUser.Put "userPrincipalName", stUserName & "@" & getDNSdomainName(stDomain)
oUser.SetPassword "Knight123"
oUser.SetInfo
iUAC = oUser.Get("userAccountControl")
If StrComp(stUserService, "service", vbTextcompare) = 0 Then
	oUser.displayName = "Application User"
	oUser.Put "userAccountControl", iUAC Or ADS_UF_DONT_EXPIRE_PASSWD Xor ADS_UF_ACCOUNTDISABLE
	oUser.SetInfo
Else
	oUser.displayName = stDispName
	oUser.pwdLastSet = 0
	oUser.Put "userAccountControl", iUAC Xor ADS_UF_ACCOUNTDISABLE
	oUser.homeDrive = "I:"
	oUser.homeDirectory = "\\" & stServer & "\" & stUserName
	oUser.scriptPath = "SLogic.bat"
	oUser.SetInfo
	Call createUserFolders(oFso, stServer, stUserName, stNBDomain)
End If
oUser.SetInfo
Call GetAccountProperties(oUser, stUserService, bInteractive)

'********************************************************************************************
Sub GetAccountProperties(byRef oUsr, Byval stUserServ, Byval bInter)
	stOut = "User has been created with the following attributes:" & vbCrlf & vbCrlf & _
	"Login ID" & vbTab & oUsr.cn & vbCrlf & _
	"First Name" & vbTab & oUsr.givenName & vbCrlf & _
	"Last Name" & vbTab & oUsr.sn & vbCrlf & _
	"Display Name" & vbTab & oUsr.displayName & vbCrlf & _
	"User Description" & vbTab & oUsr.description & vbCrlf & _
	"User OU" & vbTab & Mid(oUsr.parent, 8) & vbCrlf
	If StrComp(stUserServ, "user", vbTextcompare) = 0 Then
		stOut = stOut & "Home Drive" & vbTab & oUsr.homeDrive & vbCrlf & _
		"Home Directory" & vbTab & oUsr.homeDirectory & vbCrlf & _
		"Login Script" & vbTab & oUsr.scriptPath & vbCrlf
	End If
	stOut = stOut & "Password set to Knight123"
	If bInter Then
		MsgBox stOut, vbOkonly, stTitle
	Else
		Wscript.Echo stOut
	End If
End Sub
'********************************************************************************************

'********************************************************************************************
Sub SetAccountProperties(Byval stUserServ, Byval stTraBac, Byval stFN, Byval stMI, Byval stLN, ByRef stDesc, Byval bInt, byRef stUsrNC, byRef stUsrDispN, byRef stUsrID, byRef stSrvN)
	stUsrNC = ""
	If StrComp(stUserServ, "service", vbTextcompare) = 0 Then
		stUsrNC = "LDAP://OU=Service Accounts," & stRootNC
	Else
		stUsrNC = "LDAP://OU=" & stTraBac & ",OU=User Accounts," & stRootNC
	End If
	stUsrID = ""
	stUsrID = LCase(Left(stFN, 1) & Left(stLN, 7))
	stUsrDispN = ""
	If stMI = "" Then
		stUsrDispN = ChangeCaseToTitle(stFN) & " " & ChangeCaseToTitle(stLN)
	Else
		stUsrDispN = ChangeCaseToTitle(stFN) & " " & Left(stMI, 1) & " " & ChangeCaseToTitle(stLN)
	End If
	stSrvN = ""
	If StrComp(Left(stFN, 1), "m", vbTextcompare) <= 0 Then
		stSrvN = "JC1DFS1"
	Else
		stSrvN = "JC1DFS2"
	End If
	If bInt Then
		stUsrID = InputBox("If you would like to change User Account ID, edit field below:", stTitle, stUsrID)
		stUsrDispN = InputBox("If you would like to change User Display Name, edit field below:", stTitle, stUsrDispN)
		stSrvN = InputBox("If you would like to change Home drive server name, edit field below:", stTitle, stSrvN)
		stDesc = InputBox("If you would like to change User Account Description, edit field below:", stTitle, stDesc)
	End If
End Sub
'********************************************************************************************

'********************************************************************************************
Sub createUserFolders(byRef oFso, Byval stSrv, Byval stUsrName, Byval stNBDomain)
	stUsrFolder = "\\" & stSrv & "\users\" & stUsrName
	stUsrConfig = "\\" & stSrv & "\configuration\" & stUsrName
	If oFso.FolderExists(stUsrFolder) Then
		'Just leave it for now
	Else
		Set oUsrFolder = oFso.CreateFolder(stUsrFolder)
		Set oUsrFolder = Nothing
		Set oUsrConfig = oFso.CreateFolder(stUsrConfig)
		Set oUsrConfig = Nothing
		Set oUsrConfig = oFso.CreateFolder(stUsrConfig & "\apps")
		Set oUsrConfig = Nothing
		Set oUsrConfig = oFso.CreateFolder(stUsrConfig & "\apps\brass")
		Set oUsrConfig = Nothing
		Set oUsrConfig = oFso.CreateFolder(stUsrConfig & "\apps\mws")
		Set oUsrConfig = Nothing
		Set oUsrConfig = oFso.CreateFolder(stUsrConfig & "\apps\qms")
		Set oUsrConfig = Nothing
	'	Wscript.Sleep 3000
		Set oExec = oShell.Exec(stCurrentFolder & "\cacls " & stUsrFolder & " /T /E /G " & stNBDomain & "\" & stUsrName & ":C")
		Wscript.Echo oExec.StdOut.ReadAll
		Set oExec = oShell.Exec(stCurrentFolder & "\cacls " & stUsrConfig & " /T /E /G " & stNBDomain & "\" & stUsrName & ":C")
		Wscript.Echo oExec.StdOut.ReadAll
	End If
End Sub
'********************************************************************************************

'********************************************************************************************
Function getDNSdomainName(Byval stADSPath)
	getDNSdomainName = ""
	Set reFindDCs = New RegExp
	With reFindDCs
		.IgnoreCase = True
		.Global = True
		.Pattern = "dc\=([^,]+)"
	End With
	If reFindDCs.Test(stADSPath) Then
		Set oFoundMatches = reFindDCs.Execute(stADSPath)
		For Each oSM In oFoundMatches
			If getDNSdomainName = "" Then
				getDNSdomainName = oSM.SubMatches(0)
			Else
				getDNSdomainName = getDNSdomainName & "." & oSM.SubMatches(0)
			End If
		Next
	End If
End Function
'********************************************************************************************

'********************************************************************************************
Function GetUserType(Byval stIn)
	If InStr(1, stIn, "U", vbTextcompare) = 1 Then
		GetUserType = "User"
	Elseif InStr(1, stIn, "S", vbTextcompare) = 1 Then
		GetUserType = "Service"
	Elseif InStr(1, stIn, "T", vbTextcompare) = 1 Then
		GetUserType = "Traders"
	Elseif InStr(1, stIn, "B", vbTextcompare) = 1 Then
		GetUserType = "Backoffice"
	Else
		Call Usage()
	End If
End Function
'********************************************************************************************

'********************************************************************************************
Function ChangeCaseToTitle(stInput)
	Dim stOutput, oReg, oMatches, oMatch

	Set oReg = New RegExp
	oReg.IgnoreCase = False
	oReg.Global = True
	oReg.Pattern = "(\S+\s*)"

	stOutput = ""
	Set oMatches = oReg.Execute(stInput)
	For Each oMatch in oMatches
		stOutput = stOutput & Ucase(Left(oMatch, 1)) & LCase(Right(oMatch, Len(oMatch) - 1))
	Next
	ChangeCaseToTitle = stOutput
End Function
'********************************************************************************************:

'********************************************************************************************
Sub Usage()
	stOut = vbCrlf & Wscript.ScriptName & _
		" can use command line interface if you use arguments as below" & _
		vbCrlf & vbTab & "or run in interactive mode" & vbCrlf & vbTab & "Arguments: " & vbCrlf & _
		"#1 -- User type. Could be User or Service and can be abbreviated to U or S" & vbCrlf & _
		vbTab & "NOTE: For Service account skip next argument" & vbCrlf & _
		"#2 -- User Category. Could be Trader or Backoffice and can be abbreviated to T or B" & vbCrlf & _
		"#3 -- User First Name" & vbCrlf & "#4 -- User Middle Initial, skip if absent" & vbCrlf & _
		"#5 -- User Last Name" & vbCrlf & "#6 -- User Account Description" & vbCrlf & vbTab & "Example:" & vbCrlf & _
		Wscript.ScriptName & " U B John I Doe ""Developer in Backoffice""" & vbCrlf & _
		"to create User account in Backoffice, first name John, middle initial I, last name Doe, description Developer in Backoffice" & vbCrlf & vbTab & "Or" & vbCrlf & _
		Wscript.ScriptName & " S SQL Service ""SQL server JC1WSSQL10 service account""" & vbCrlf & _
		"to create Service account, first name equivalent is SQL, last name equivalent is Service, description SQL server JC1WSSQL10 service account" & vbCrlf
	If bInteractive Then
		MsgBox stOut, vbOkonly, stTitle
	Else
		Wscript.Echo stOut
	End If
	Wscript.Quit
End Sub

'********************************************************************************************
