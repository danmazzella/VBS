On Error Resume Next
Dim fso, folder, files, NewsFile,sFolder

Set fso = CreateObject("Scripting.FileSystemObject")

 msgbox("This program requires an input from command line with individual files named as the users login name")

'Takes input from command line
sFolder = Wscript.Arguments.Item(0)
Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Copy AIM Folder Users I Drive"""
msgbox("Script Complete")

If sFolder = "" Then
  Wscript.Echo "No Folder parameter was passed"
  Wscript.Quit
End If
'Get Folder and Files from Input
Set folder = fso.GetFolder(sFolder)
Set files = folder.Files

'For every 
For each filefName In files
	'Deletes the .xlsx from file
	fName = Left(FSO.GetFilefName(filefName), Len(FSO.GetFilefName(filefName)) - 5)
	'Windows Logon limited to 8 characters
	if Len(fName) > 8 then
		fName = Left(fName, Len(fName) - (Len(fName) - 8))
	end if	
	
	'Set directory to JC1DFS2
	strDirectory = "\\jc1dfs2\users\" & fName
	'If the users I: drive exists then
	IF FSO.FolderExists(strDirectory) THEN
		'Create AIM PASS folder
		strDirectory = "\\jc1dfs2\users\" & fName & "\AIM Pass"
		Set objFolder = FSO.CreateFolder(strDirectory)
		WScript.Echo "Just created " & strDirectory 
		'Copy the file to the I:AIM\Pass
		FSO.CopyFile "C:\temp\AIM Password\" & FSO.GetFilefName(filefName), "\\jc1dfs2\users\" & fName & "\AIM Pass\" & FSO.GetFilefName(filefName)
	ELSE
		'If not on DFS2, then DFS1
		strDirectory = "\\jc1dfs1\users\" & fName
		'If users I drive exsists:
		if FSO.FolderExists(strDirectory) THEN
			'Create Folder
			strDirectory = "\\jc1dfs1\users\" & fName & "\AIM Pass"
			Set objFolder = FSO.CreateFolder(strDirectory)
			WScript.Echo "Just created " & strDirectory 
			'Copy the file
			FSO.CopyFile "C:\temp\AIM Password\" & FSO.GetFilefName(filefName), "\\jc1dfs1\users\" & fName & "\AIM Pass\" & FSO.GetFilefName(filefName)
		else
			'If they don't have an I drive then write to the .txt file
			Set WriteStuff = FSO.OpenTextFile("c:\temp\AIM Password\Fail.txt", 8, True)
			WriteStuff.WriteLine(fName)
			WriteStuff.Close
			SET WriteStuff = NOTHING
		end if
	END IF
	wscript.echo VbCrLf	
Next 