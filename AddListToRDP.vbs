option explicit
msgbox "This program requires a file named computernames.txt listing all desired computers." & VbCrLf & "Located in C:\Temp"
ReadTXT
Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Add List of Users to Local RDP/Admin"""
MsgBox("Script Complete")
wscript.quit
 
 sub ReadTXT()
	dim FSO, objTextFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = FSO.OpenTextFile("c:\temp\computernames.txt")
		Do Until objTextFile.AtEndOfStream
			AddtoGroup objTextFile.Readline
		Loop
  end sub

 sub AddtoGroup(computerName)
	dim msg, objLocalGroup, strInput, objDomainUser, objPing, objStatus
	msg = computerName

	Set objLocalGroup = GetObject("WinNT://" & msg & "/Remote Desktop Users,group")
	wscript.echo "The desired PC is:" & msg
	
	'Input Box/Name
	strInput = InputBox( "Enter the desired username:" )
    Wscript.Echo "The desired username is: " & strInput

	Set objDomainUser = GetObject("WinNT://Knight/" & strInput & ",user")
	
	'Ping
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._  
	ExecQuery("select Replysize from Win32_PingStatus where address = '" & msg & "'")  
	
	For Each objStatus in objPing  
		If  IsNull(objStatus.ReplySize) Then  
			msgBox "Computer is offline, user not added"  
		Else  
			WScript.Echo "Computer is online"  
			
			'If user invalid
			on error resume next 
			Set objDomainUser = GetObject("WinNT://Knight/dmazzell,user")
			if ( err ) then 
				msgBox "User does not exist"
			else 
				'If Member Exists	
				If (objLocalGroup.IsMember(objDomainUser.ADsPath) = False) Then
					objLocalGroup.Add(objDomainUser.ADsPath)
				Else
					msgBox "User already exists in Remote Desktop"
				End If
			End if 		
		End If
	Next  
  
	Set objPing=Nothing  
	Set objStatus=Nothing  
 end sub