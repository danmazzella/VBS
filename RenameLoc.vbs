dim message

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

Set args = WScript.Arguments
newComputerName = args.Item(0)

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
	objComputer.JoinDomainOrWorkgroup "knight.global.com", "dmazzelladm", "Danny2010"
Next

If ErrCode = 0 Then
	wscript.echo "0"
Else 
	wscript.echo errcode
End If 
	
	
	
	
