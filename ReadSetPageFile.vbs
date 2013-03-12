'What group to add to
PageComputer = "0"
do while PageComputer <> "1" and PageComputer <> "2"
	PageComputer = InputBox("1 - Read Page File" &  VbCrLf & "2 - Set Page File" & VbCrLf & VbCrLf & "Would you like to read page file or set page file? ")
	If PageComputer = False Then
		MsgBox "You pressed cancel!"
		wscript.quit
End If
loop

'Enter PC Name
strComputer = InputBox( "Enter PC: " )

'Ping (Make sure computer is pingable)
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._  
ExecQuery("select Replysize from Win32_PingStatus where address = '" & strComputer & "'")  

For Each objStatus in objPing  
	'Computer does not exist
	If  IsNull(objStatus.ReplySize) Then  
		msgBox "Computer is offline, No user added"  
	Else  
		if PageComputer = "1" then 
			ReadPageFile(strComputer)
		elseif PageComputer = "2" then
			SetPageFile(strComputer)
		End if
	end if
Next

sub ReadPageFile(strComputer)
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colPageFiles = objWMIService.ExecQuery _
		("Select * from Win32_PageFile")
	For each objPageFile in colPageFiles
		msgBox(	"PC Name: " & vbTab & strComputer & VbCrLf & _
				"CreationDate: " & vbTab &  objPageFile.CreationDate & VbCrLf & _
				"Description: " & vbTab &  objPageFile.Description & VbCrLf & _
				"Drive: " & vbTab & vbTab &  objPageFile.Drive & VbCrLf & _
				"FileName: " & vbTab &  objPageFile.FileName & VbCrLf & _ 
				"FileSize: " & vbTab & vbTab &  objPageFile.FileSize & VbCrLf & _
				"InitialSize: " & vbTab & vbTab &  objPageFile.InitialSize & VbCrLf & _
				"InstallDate: " & vbTab &  objPageFile.InstallDate & VbCrLf & _
				"MaximumSize: " & vbTab &  objPageFile.MaximumSize & VbCrLf & _
				"Name: " & vbTab & vbTab &  objPageFile.Name & VbCrLf & _
				"Path: " & vbTab & vbTab &  objPageFile.Path)
	Next
end sub

sub SetPageFile(strComputer)
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colPageFiles = objWMIService.ExecQuery _
		("Select * from Win32_PageFileSetting")
	For Each objPageFile in colPageFiles
		objPageFile.InitialSize = 2046
		objPageFile.MaximumSize = 4092
		objPageFile.Put_
	Next
end sub


Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Read or Set Page File"""
msgbox("Script Complete")