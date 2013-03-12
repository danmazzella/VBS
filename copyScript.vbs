If WScript.Arguments.length =0 Then
  Set objShell = CreateObject("Shell.Application")
  'Pass a bogus argument with leading blank space, say [ uac]
  objShell.ShellExecute "wscript.exe", Chr(34) & _
  WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
Else
    Set objNetwork = CreateObject("WScript.Network")
	SET objShell = CREATEOBJECT("Wscript.Shell")
	UserName = objNetwork.UserName

	'Grabs adm
	fName = Right(UserName, 3)

	if fName = "adm" then
		objShell.Run "\\jc1dfs2\systems\Script.bat"
	else
		strUser = InputBox("Enter username")
		objShell.Run "runas /profile /user:knight\" & strUser & " \\jc1dfs2\systems\Script.bat"
	end if
End If



