Set fso = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = FSO.OpenTextFile("\\jc1wsalt03\express\temp\DanLog\Log.csv", 8, True)
Set objNTInfo = CreateObject("WinNTSystemInfo")
Line = LCase(WScript.Arguments.Item(0))
If Line = "e" then
	script = "Explorer++"
elseif Line = "d" then
	script = "Device Manager"
elseif Line = "c" then
	script = "Command Prompt"
elseif Line = "r" then
	script = "Rename computer"
elseif Line = "s" then
	script = "Shutdown Computer"
elseif Line = "b" then
	script = "Restart Computer"
elseif Line = "i" then
	script = "Internet Explorer"
elseif Line = "p" then
	script = "System Properties"
elseif Line = "m" then
	script = "Computer Management"
elseif Line = "a" then
	script = "App/Remove Programs"
elseif Line = "q" then
	script = "Quit"	
else
	script = Line & " ??"
end if
WriteStuff.WriteLine(Date & "," & Time & "," & objNTInfo.ComputerName & "," & objNTInfo.UserName & "," & "Copy.bat - " & script)
WriteStuff.Close