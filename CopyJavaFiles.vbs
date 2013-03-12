Dim fso, folder, files, sFolder
public exists, dexist, wrongSize
exists = "0"
dexist = "0"
wrongSize = "0"
Set fso = CreateObject("Scripting.FileSystemObject")

strComputer = InputBox("Enter PC NAME: ")

ComputerOS = OperatingSystem(strComputer)

if ComputerOS = "Win7" then
	sfolder = "\\jc1wsalt03\Library\x64-Packages\sun\JavaRE\1.6.20 x64\justfiles\java\jre1620"
elseif ComputerOS = "XP" then
	sfolder = "\\jc1wsalt03\library\Packages\Sun\Java\1.6.0.20\justfiles\java\jre1620"
end if

strDirectory = "\\" & strComputer & "\C$\Program Files\Java"

'Get Folder and Files from Input
Set folder = fso.GetFolder(sFolder)
Set files = folder.Files
subfolder sfolder, strDirectory, counter
wscript.echo exists & " already existed and " & dexist & " did not exist " & wrongSize & " were the wrong size"
MSGBox(exists & " already existed and " & dexist & " did not exist " & wrongSize & " were the wrong size")
Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Copy Missing Java 1.6.0.20 Files"""
msgbox("Script Complete")
sub subfolder(strSource, strDestination, counter)
	'Connect to the current directory in strSource. 
	Set objDir = FSO.GetFolder(strSource) 

	'If destination folder doesn't exist, create it. 
	If Not FSO.FolderExists(strDestination) Then 
		FSO.CreateFolder(strDestination) 
	End If 

	'If current folder doesn't exist under destination folder, create it. 
	If Not FSO.FolderExists(strDestination & "\" & objDir.Name) Then 
		FSO.CreateFolder(strDestination & "\" & objDir.Name) 
	End If 

	For Each objFiles In FSO.GetFolder(strSource).Files 
		counter = counter + 1
		FileThere objFiles, strDestination & "\" & objDir.Name
	Next
	For Each objFolder In FSO.GetFolder(strSource).SubFolders 
			subfolder objFolder.Path, strDestination & "\" & objDir.Name, counter
	Next 
end sub

sub FileThere(filefName, strDirectory)
	IF FSO.FileExists(strDirectory & "\" & FSO.GetFileName(filefName)) Then
		Set objFile = FSO.GetFile(strDirectory & "\" & FSO.GetFileName(filefName))
		if filefName.Size = objFile.Size THEN
			wscript.echo "Exists: " & FSO.GetFileName(filefName)
			exists = exists + 1
		ELSEIF filefName.Size <> objFile.Size THEN
			wscript.echo "Wrong File Size: " & FSO.GetFileName(filefName)
			FSO.DeleteFile strDirectory & "\" & FSO.GetFileName(filefName)
			FSO.CopyFile filefName, strDirectory & "\" & jre1620
			wrongSize = wrongSize + 1
		end if
	ELSEIF NOT FSO.FileExists(strDirectory & "\" & FSO.GetFileName(filefName)) THEN
		wscript.echo "Doesn't exist: " & FSO.GetFileName(filefName)
		FSO.CopyFile filefName, strDirectory & "\" & jre1620
		dexist = dexist + 1
	END IF
end sub

function OperatingSystem(Comps)
	'OS Verification
	strComputer = Comps
	Set oWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colOSInfo = oWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each oOSProperty in colOSInfo 
		strCaption = oOSProperty.Caption 
	Next

	If InStr(1,strCaption, "Windows 7", vbTextCompare) Then
		OperatingSystem = "Win7"
	end if
	If InStr(1,strCaption, "XP", vbTextCompare) Then 
		OperatingSystem = "XP"
	end if
end function