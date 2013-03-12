dim message

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

message = "Please enter computer name. Leave blank or press cancel to quit. "
newComputerName = InputBox(message, title)

Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
    strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & _
        strComputer & "'")
		

If newComputerName = "" Then
    Wscript.quit
End If

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colComputers = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")

For Each objComputer in colComputers
	Errcode = objComputer.Rename(newComputerName)
Next

If ErrCode = 0 Then
	MsgBox "Computer renamed correctly."
Else 
	MsgBox "Computer Name not Changed."
	msgbox errcode
End If 
	
	
	
	
