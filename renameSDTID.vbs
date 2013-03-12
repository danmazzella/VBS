Continue = MsgBox("This will replace all files in RSA folder." & VbCrLf & "Would you like to continue?", vbYesNo, "Warning!")

if Continue = 6 then
	strPath = "c:\users\dmazzell\desktop\RSA\"

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set FLD = FSO.GetFolder(strPath)


	For Each daFil in FLD.Files
		strOldName = daFil.Name
		strOldPath = daFil.Path
		
		If InStr(strOldName, "sdtid") > 0 Then
			If InStr(strOldName, "_") > 0 Then
				strFileParts = Split(LCASE(strOldName), "_")

				dim objTextFile
				Set objTextFile = FSO.OpenTextFile("c:\users\dmazzell\desktop\AD.csv")
				Do Until objTextFile.AtEndOfStream
					ADinfo = objTextFile.Readline
					ADsplit = Split(LCASE(ADinfo), ",")
					if strFileParts(0) = ADsplit(0) then
						'msgbox strFileParts(0) & " - " & ADsplit(1)
						strNewName = ADsplit(1)
						
						if strNewName <> "" then
							if NOT FSO.FileExists(strPath & strNewName & ".sdtid") then
								FSO.MoveFile strPath & strOldName, strPath & strNewName & ".sdtid"
							else
								if NOT FSO.FileExists(strPath & strNewName & "2.sdtid") then
									FSO.MoveFile strPath & strOldName, strPath & strNewName & "2.sdtid"
								else
									if NOT FSO.FileExists(strPath & strNewName & "3.sdtid") then
										FSO.MoveFile strPath & strOldName, strPath & strNewName & "3.sdtid"
									else
										if NOT FSO.FileExists(strPath & strNewName & "4.sdtid") then
											FSO.MoveFile strPath & strOldName, strPath & strNewName & "4.sdtid"
										end if
									end if
								end if
							end if
							Exit Do
						End if
					end if
				Loop
			End If
		End If
	Next

	Set FLD = Nothing
	Set FSO = Nothing
End If