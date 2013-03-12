
If WScript.Arguments.length =0 Then
  Set objShell = CreateObject("Shell.Application")
  'Pass a bogus argument with leading blank space, say [ uac]
  objShell.ShellExecute "wscript.exe", Chr(34) & _
  WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
Else
    strComputer = InputBox("Fix Remedy for Which PC?","PC Name?",strComputer)
	
	Const HKEY_LOCAL_MACHINE = &H80000002
	 
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
		strComputer & "\root\default:StdRegProv")
	 
	strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\CB1C5AA8F28A89340B497B487D1D7D3C"
	strValueName = "E07488A54E962C54EA3634BC5C61ABD1"
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath, _
		strValueName, "I:\Remedy\AR"

	strValueName = "10000000000000000000000000000000"
	oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath, _
		strValueName, "I:\Remedy\AR"
		

	Dim objShell
	SET objShell = CREATEOBJECT("Wscript.Shell")
	objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Fix Remedy"""
	
	msgbox "Please try to relaunch Remedy"
End If
