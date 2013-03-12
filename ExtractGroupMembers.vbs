Set fso = CreateObject("Scripting.FileSystemObject")
Set WriteStuff = FSO.OpenTextFile("c:\users\dmazzell\desktop\users.txt", 8, True)
MsgBox "This script exports all users of an AD group, if you want to change group contact Dan Mazzella"
Set objGroup = GetObject("LDAP://cn=brassapache,ou=UNIX Groups,ou=Groups,dc=global,dc=knight,dc=com")

For Each objMember in objGroup.Members
	WriteStuff.WriteLine(objMember.Name)
Next

Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Export Members of An AD Group"""