Set fso = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = FSO.OpenTextFile("\\jc1wsalt03\express\temp\DanLog\Log2.csv", 8, True)
Set objNTInfo = CreateObject("WinNTSystemInfo")
Line = WScript.Arguments.Item(0)

WriteStuff.WriteLine(Date & "," & Time & "," & objNTInfo.ComputerName & "," & objNTInfo.UserName & "," & Line)
WriteStuff.Close