Set objNetwork = CreateObject("WScript.Network")
UserName = objNetwork.UserName

strComputer = "."
Set oWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colOSInfo = oWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each oOSProperty in colOSInfo 
	strCaption = oOSProperty.Caption 
Next

If InStr(1,strCaption, "Windows 7", vbTextCompare) Then
	If WScript.Arguments.length =0 Then
		Set oShell = CreateObject("Shell.Application")
		'Pass a bogus argument with leading blank space, say [ uac]
		oShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
	else
		SET objShell = CREATEOBJECT("Wscript.Shell")
		'Grabs adm
		fName = Right(UserName, 3)	
		if fName = "adm" then
			objShell.Run "\\jc1dfs2\applications\desktop\CopyBat\Script.bat"
		else
			strUser = InputBox("Enter username")
			objShell.Run "runas /profile /user:knight\" & strUser & " \\jc1dfs2\applications\desktop\copybat\Script.bat"
		end if
	end if
elseif InStr(1,strCaption, "XP", vbTextCompare) Then 
	SET objShell = CREATEOBJECT("Wscript.Shell")
	'Grabs adm
	fName = Right(UserName, 3)	
	if fName = "adm" then
		objShell.Run "\\jc1dfs2\applications\desktop\CopyBat\Script.bat"
	else
		strUser = InputBox("Enter username")
		objShell.Run "runas /profile /user:knight\" & strUser & " \\jc1dfs2\applications\desktop\copybat\Script.bat"
	end if
end if	