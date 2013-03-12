Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const xlLine = 4, xlColumns = 2, xlLocationAsNewSheet = 1, xlCategory = 1, xlPrimary = 1, xlValue = 2, xlThin = 2
Const xlNone = &HFFFFEFD2, xlCategoryScale = 2, xlWorkbookNormal = &HFFFFEFD1
Dim arCPU(), arNetwork(), arMemory(), arLogs()
Dim arCPUavg(), arNetworkAvg(), arMemoryAvg()
Dim arCPUHeader(3), arNetworkHeader(), arMemHeader()
'	Creating file header hash and defining necessary arrays:
Set dicProcHeader = Wscript.CreateObject("Scripting.Dictionary")
Dim arProcCPU(), arProcCPUtime(), arProcMemK(), arProcPageFaults()
Dim arProcCPUAvg(), arProcCPUtimeAvg(), arProcMemKAvg(), arProcPageFaultsAvg()
Dim arProcCPUMax(), arProcCPUtimeMax(), arProcMemKMax(), arProcPageFaultsMax()
Redim arProcCPUAvg(1), arProcCPUtimeAvg(1), arProcMemKAvg(1), arProcPageFaultsAvg(1)
Redim arProcCPUMax(1), arProcCPUtimeMax(1), arProcMemKMax(1), arProcPageFaultsMax(1)
Redim arProcCPU(0), arProcCPUtime(0), arProcMemK(0), arProcPageFaults(0)
'On Error Resume Next
Dim arHead()
stTitle = "Processing KTop Log files"
stPrompt1 = "Total number of files, available for processing: "
stPrompt2 = "Please, specify which file you would like to process: "
stPrompt3 = "Would you like to calculate per-process statistics (longer)?"
stPrompt4 = "To minimize clutter report will be presented for top N processes." & vbCrlf & _
	"Please, specify number of processes to report:"
stStart = Timer
Set oArgs = Wscript.Arguments
If oArgs.Count = 1 Then
	nSecs = CInt(oArgs(0))
Else
	nSecs = CInt(300)
End If

'This is to avoid comparison with UBound of uninitiated array
Redim arCPU(3)
Redim arNetwork(0)
Redim arMemory(0)

Set oFso = Wscript.CreateObject("Scripting.FileSystemObject")

If Not oFso.FolderExists("c:\temp") Then oFso.CreateFolder("c:\temp")
If Not oFso.FolderExists("C:\temp\DanScripts\KTOP\") Then oFso.CreateFolder("C:\temp\DanScripts\KTOP\")
stBaseDir = "C:\temp\DanScripts\KTOP\"

stSourceDir = stBaseDir
'stSourceDir = "c:\KnightLogs\Ktop\"
Set oSourceDir = oFso.GetFolder(stSourceDir)
Set reLogName = New RegExp
With reLogName
	.Global = False
	.IgnoreCase = False
	.Pattern = "^([0-9]{8}-[^.-]+)\.txt$"
End With
dtLatest = DateAdd("d", -70, Now())
iLogs = 0
For Each oFile In oSourceDir.Files
	If reLogName.Test(oFile.Name) Then
		Set oFoundLogs = reLogName.Execute(oFile.Name)
		If oFso.FileExists(stBaseDir & oFoundLogs(0).SubMatches(0) & ".xls") Then
			'File already processed, skipping
		Else
			Redim Preserve arLogs(iLogs)
			arLogs(iLogs) = oFoundLogs(0).SubMatches(0)
		'	Wscript.Echo "Found file: " & oFoundLogs(0).SubMatches(0)
			iLogs = iLogs + 1
		End If
	End If
Next
Set reProcs = New RegExp
With reProcs
	.Global = True
	.IgnoreCase = True
	.Pattern = "(^\d{2}:\d{2}:\d{2})\s+(\d+)\s+([\w\.]+)\s+([\d\.]+)%\s+(\d{1,2}:\d{2}:\d{2})\s+(\d+)k\s+" & _
		"(\d+)\s+(\d+)k\s+(\w+)\s+(\d+)\s+(.*)"
End With
Set reHeader = New RegExp
With reHeader
	.Global = False
	.IgnoreCase = False
	.Pattern = "(^[0-9]{2}:[0-9]{2}:[0-9]{2})\s([^:]+):\s+([\d\.]+)"
End With
Set reSpace = New RegExp
With reSpace
	.IgnoreCase = True
	.Global = True
	.Pattern = "(\s+)"
End With
Set reCPU = New RegExp
With reCPU
	.Global = False
	.IgnoreCase = False
	.Pattern = "(^\d{2}:\d{2}:\d{2})\s([^:]+):\s[^:]+:\s+([\d\.]+)%[-\s(]+(\w+):\s+([\d\.]+)%,\s+(\w+):\s+([\d\.]+)%\)"
	'.Pattern = "(^\d{2}:\d{2}:\d{2})\s([^:]+):\s+([\d\.]+)%[-\s\(]+(\w+):\s+([\.\d]+)%,\s+(\w+):\s+([\.\d]+)%\)"
End With
Set reCPU_old = New RegExp
With reCPU_old
	.Global = False
	.IgnoreCase = False
	.Pattern = "(^\d{2}:\d{2}:\d{2})\s([^:]+):\s+([\d\.]+)%[-\s\(]+(\w+):\s+([\.\d]+)%,\s+(\w+):\s+([\.\d]+)%\)"
End With

nTopN = 15
Set oExcel = WScript.CreateObject("Excel.Application")
oExcel.Application.DisplayAlerts = False
oExcel.SheetsInNewWorkbook = 1
'oExcel.Visible = True
oExcel.Visible = False

For Each stDateCompName In arLogs
	'stDateCompName = Mid(stKTopLog, 1, InStr(stKTopLog, ".") - 1)
	stKTopLog = stDateCompName & ".txt"
	stCPU = stBaseDir & stDateCompName & "-CPU.csv"
	stMemory = stBaseDir & stDateCompName & "-Memory.csv"
	stNetwork = stBaseDir & stDateCompName & "-Network.csv"
	stProcess = stBaseDir & stDateCompName & "-Process.csv"
	''stProcAvgCPU = stBaseDir & stDateCompName & "-ProcAverage-CPU.csv"
	'stProcAvgCPUtime = stBaseDir & stDateCompName & "-ProcAverage-CPUtime.csv"
	''stProcAvgMemK = stBaseDir & stDateCompName & "-ProcAverage-MemK.csv"
	''stProcAvgPageFaults = stBaseDir & stDateCompName & "-ProcAverage-PF.csv"
	''Set tsProcAvgCPU = oFso.OpenTextFile(stProcAvgCPU, ForWriting, True)
	'Set tsProcAvgCPUtime = oFso.OpenTextFile(stProcAvgCPUtime, ForWriting, True)
	''Set tsProcAvgMemK = oFso.OpenTextFile(stProcAvgMemK, ForWriting, True)
	''Set tsProcAvgPageFaults = oFso.OpenTextFile(stProcAvgPageFaults, ForWriting, True)
	Wscript.Echo "Source Logs File: " & stSourceDir & stKTopLog
	Wscript.Echo "Converting File....Please Wait"
	Set tsKTopLog = oFso.OpenTextFile(stSourceDir & stKTopLog, ForReading, False)
	Set tsNetwork = oFso.OpenTextFile(stNetwork, ForWriting, True)
	Set tsMemory = oFso.OpenTextFile(stMemory, ForWriting, True)
	Set tsCPU = oFso.OpenTextFile(stCPU, ForWriting, True)
	Set tsProcess = oFso.OpenTextFile(stProcess, ForWriting, True)
	bFirst = True
	bFirstProc = True
	bNotFlushed = True
	bStartNet = False
	bStartCPU = False
	bStartMem = False
	bStartProc = False
	iProc = 1
	iTimeCount = CInt(-1)
	iAvgTimeCount = CInt(0)
	iAvgNumber = 1
	stProcCPUout = ""
	stProcMemKout = ""
	stProcPageFaultsOut = ""
	'stKTopAll = tsKTopLog.ReadAll
	'stProcessDetails = stKTopAll
	'arKTopAll = Split(stKTopAll, vbCrlf, -1, vbBinarycompare)
	'For Each stLine In arKTopAll
	Do While Not tsKTopLog.AtEndOfStream
	'For i = 0 To 1000
		stLine = ""
		stLine = tsKTopLog.ReadLine
		If InStr(1, stLine, "-- Knight Top", vbTextcompare) Then
		'	Wscript.Echo vbTab & "Header found"
			If bFirst Then
				'Header at the beginning of the file -- skip that case
				bFirst = False
			Else
				If bNotFlushed Then
					tsNetwork.WriteLine Join(arNetworkHeader, ",")
					tsCPU.WriteLine Join(arCPUHeader, ",")
					tsMemory.WriteLine Join(arMemHeader, ",")
					dtStart = arCPU(0)
					dtStartProc = arCPU(0)
					bNotFlushed = False
				End If
				If nSecs < 5 Then
					tsNetwork.WriteLine Join(arNetwork, ",")
					tsCPU.WriteLine Join(arCPU, ",")
					tsMemory.WriteLine Join(arMemory, ",")
					iAvgTimeCount = iAvgTimeCount + 1
				Else
				'	Wscript.Echo "dtStart = " & dtStart & vbTab & "arCPU(0) = " & arCPU(0) & vbTab & _
				'		"DateDiff = " & DateDiff("s", dtStart, arCPU(0)) & vbTab & _
				'		"Comparizon: " & CStr(DateDiff("s", dtStart, arCPU(0)) > CInt(nSecs))
					If DateDiff("s", dtStart, arCPU(0)) > CInt(nSecs) Then
						bAveraged = True
						dtStart = arCPU(0)
					Else
						bAveraged = False
					End If
					If bAveraged Then
					'	stNow = Timer
				'		Wscript.Echo "Average processing: " & arCPU(0)' & " took " & CSng(stNow) - CSng(stPrev) & " seconds"
					'	stPrev = Timer
						nAvgNumber = iAvgNumber
						Call DivideArray(arNetworkAvg, nAvgNumber, 1)
						Call DivideArray(arCPUavg, nAvgNumber, 1)
						Call DivideArray(arMemoryAvg, nAvgNumber, 1)
						tsNetwork.WriteLine Join(arNetworkAvg, ",")
						tsCPU.WriteLine Join(arCPUavg, ",")
						tsMemory.WriteLine Join(arMemoryAvg, ",")
						iAvgTimeCount = iAvgTimeCount + 1
						Call CopyArray(arNetwork, arNetworkAvg, 1)
						Call CopyArray(arCPU, arCPUavg, 1)
						Call CopyArray(arMemory, arMemoryAvg, 1)
						iAvgNumber = 1
					Else
					'	Wscript.Echo " Regular processing: " & arCPU(0)
						Redim Preserve arCPUavg(UBound(arCPU))
						Redim Preserve arNetworkAvg(UBound(arNetwork))
						Redim Preserve arMemoryAvg(UBound(arMemory))
						arCPUavg(0) = arCPU(0)
						arNetworkAvg(0) = arNetwork(0)
						arMemoryAvg(0) = arMemory(0)
						Call CopyAvgArray(arNetwork, arNetworkAvg, 1)
						Call CopyAvgArray(arCPU, arCPUavg, 1)
						Call CopyAvgArray(arMemory, arMemoryAvg, 1)
						Call ZeroArray(arNetwork, 1)
						Call ZeroArray(arCPU, 1)
						Call ZeroArray(arMemory, 1)
						iAvgNumber = iAvgNumber + 1
					End If
				End If
				iTimeCount = iTimeCount + 1
			End If
		Elseif InStr(1, stLine, "Network Statistics:", vbBinarycompare) Then
			bStartNet = True
			iNetInd = 1
		Elseif InStr(1, stLine, "Memory Statistics:", vbBinarycompare) Then
			bStartMem = True
			iMemInd = 1
		Elseif InStr(1, stLine, "CPU Statistics:", vbBinarycompare) Then
			bStartCPU = True
			iCPUInd = 1
		Elseif InStr(1, stLine, "Usage: CPU_Total", vbBinarycompare) Then
			'Wscript.Echo vbTab & "processing " & stLine
			If reCPU.Test(stLine) Then
				Set oCPUParms = reCPU.Execute(stLine)
				arCPUHeader(0) = "Time"
				arCPUHeader(1) = "CPU Usage Total"
				arCPUHeader(2) = "CPU Usage Kernel"
				arCPUHeader(3) = "CPU Usage User"
				arCPU(0) = oCPUParms(0).SubMatches(0)
				arCPU(1) = oCPUParms(0).SubMatches(2)
				arCPU(2) = oCPUParms(0).SubMatches(4)
				arCPU(3) = oCPUParms(0).SubMatches(6)
			End If
		Elseif InStr(1, stLine, "Process Statistics:", vbBinarycompare) Then
			bStartProc = True
			iProcInd = 1
		Elseif stLine = "" Then
			bStartNet = False
			bStartCPU = False
			bStartMem = False
			bStartProc = False
		Else
		'	Wscript.Echo stLine
			If reHeader.Test(stLine) Then
				Set oDataLine = reHeader.Execute(stLine)
				If bStartNet Then
					If iNetInd > UBound(arNetwork) Then
						Redim Preserve arNetwork(iNetInd)
						Redim Preserve arNetworkHeader(iNetInd)
						arNetworkHeader(0) = "Time"
					End If
					If Ubound(Filter(arNetworkHeader, oDataLine(0).SubMatches(1), True, vbTextcompare)) < 0 Then _
							arNetworkHeader(iNetInd) = oDataLine(0).SubMatches(1)
					arNetwork(0) = oDataLine(0).SubMatches(0)
					arNetwork(iNetInd) = oDataLine(0).SubMatches(2)
					iNetInd = iNetInd + 1
				Elseif bStartCPU Then
					'Wscript.Echo "CPU calculation: " & reCPU.Test(stLine)
					If reCPU_old.Test(stLine) Then
						Set oCPUParms = reCPU_old.Execute(stLine)
						arCPUHeader(0) = "Time"
						arCPUHeader(1) = "CPU Usage Total"
						arCPUHeader(2) = "CPU Usage Kernel"
						arCPUHeader(3) = "CPU Usage User"
						arCPU(0) = oCPUParms(0).SubMatches(0)
						arCPU(1) = oCPUParms(0).SubMatches(2)
						arCPU(2) = oCPUParms(0).SubMatches(4)
						arCPU(3) = oCPUParms(0).SubMatches(6)
					End If
				Elseif bStartMem Then
					If iMemInd > UBound(arMemory) Then
						Redim Preserve arMemory(iMemInd)
						Redim Preserve arMemHeader(iMemInd)
						arMemHeader(0) = "Time"
					End If
					If Ubound(Filter(arMemHeader, oDataLine(0).SubMatches(1), True, vbTextcompare)) < 0 Then
						arMemHeader(iMemInd) = oDataLine(0).SubMatches(1)
					Else
						arTemp = Filter(arMemHeader, oDataLine(0).SubMatches(1), True, vbTextcompare)
						If arTemp(0) <> oDataLine(0).SubMatches(1) Then arMemHeader(iMemInd) = oDataLine(0).SubMatches(1)
					End If
					arMemory(0) = oDataLine(0).SubMatches(0)
					arMemory(iMemInd) = oDataLine(0).SubMatches(2)
					iMemInd = iMemInd + 1
				End If
			Else
				If bStartProc Then
					If reProcs.Test(stLine) Then
						Set oFoundParams = reProcs.Execute(stLine)
						stTemp = ""
						For Each stParam In oFoundParams(0).SubMatches
							If stTemp = "" Then
								stTemp = stParam
							Else
								stTemp = stTemp & "," & stParam
							End If
						Next
					'	stProcessDetails = stProcessDetails & vbCrlf & stTemp
						Call ProcDetails(stTemp)
				'		tsProcess.WriteLine stTemp
					End If
				End If
			End If
		End If
	Loop
	'Next
	tsKTopLog.Close
	tsNetwork.Close
	tsMemory.Close
	tsCPU.Close
	Set tsKTopLog = Nothing
	Set tsNetwork = Nothing
	Set tsMemory = Nothing
	Set tsCPU = Nothing
	
	tsProcess.Close
	'stMid = Timer
	'Wscript.Echo "Processed summaries: " & CSng(stMid) - CSng(stStart) & " seconds."
	'	Getting indexes of colums, sorted by column maximum for "top N" report
	'	and retaining only nTopN elements in every Max array. If number of Top N is bigger than
	'	total number of processes, resetting Top N to latter
	If CInt(nTopN) > CInt(UBound(dicProcHeader.Keys) + 1) Then nTopN = CInt(UBound(dicProcHeader.Keys) + 1)
	'Wscript.Echo "arProcCPUMax = (" & Join(arProcCPUMax, ",") & ")"
	Call SortArrayInd(arProcCPUMax, 1, False, nTopN)
	'Wscript.Echo "After: arProcCPUMax = (" & Join(arProcCPUMax, ",") & ")"
	'Wscript.Echo "arProcPageFaultsMax = (" & Join(arProcPageFaultsMax, ",") & ")"
	Call SortArrayInd(arProcPageFaultsMax, 1, False, nTopN)
	'Wscript.Echo "After: arProcPageFaultsMax = (" & Join(arProcPageFaultsMax, ",") & ")"
	'Wscript.Echo "arProcMemKMax = (" & Join(arProcMemKMax, ",") & ")"
	Call SortArrayInd(arProcMemKMax, 1, False, nTopN)
	'Wscript.Echo "After: arProcMemKMax = (" & Join(arProcMemKMax, ",") & ")"
	'	Create headers
	Redim arHead(UBound(dicProcHeader.Keys) + 1)
	For Each stKey In dicProcHeader.Keys
		'Wscript.Echo "Key: " & stKey & vbTab & "Element number: " & dicProcHeader(stKey)
		arHead(dicProcHeader(stKey)) = stKey
	Next
	arHead(0) = "Time"
	'	Appending headers to output strings
	stProcCPUout = Join(arHead, ",") & vbCrlf & stProcCPUout
	stProcPageFaultsOut = Join(arHead, ",") & vbCrlf & stProcPageFaultsOut
	stMemHead = "Time"
	For iMem = 1 To UBound(arHead)
		stMemHead = stMemHead & "," & arHead(iMem) & "-Memory~" & arHead(iMem) & "-Virt.Memory"
	Next
	arMemHead = Split(Replace(stMemHead, "~", ",", 1, -1, vbBinarycompare), ",", -1, vbBinarycompare)
	stProcMemKout = stMemHead & vbCrlf & stProcMemKout
	''	Dumping averaged results for all processes to file for analysis:
	'tsProcAvgCPU.WriteLine stProcCPUout
	'tsProcAvgPageFaults.WriteLine stProcPageFaultsOut
	'tsProcAvgMemK.WriteLine Replace(stProcMemKout, "~", ",", 1, -1, vbBinarycompare)
	''	tsProcAvgMemK.WriteLine stProcMemKout
	'tsProcAvgCPU.Close
	'tsProcAvgPageFaults.Close
	'tsProcAvgMemK.Close
	'Set tsProcAvgCPU = Nothing
	'Set tsProcAvgPageFaults = Nothing
	'Set tsProcAvgMemK = Nothing
	
	'	Re-assembling output strings to include only Top N processes into report
	Call GetTopN(stProcCPUout, arProcCPUMax)
	Call GetTopN(stProcPageFaultsOut, arProcPageFaultsMax)
	Call GetTopN(stProcMemKout, arProcMemKMax)
	
	stProcAvgCPUTop = stBaseDir & stDateCompName & "-ProcAverage-CPU-Top" & nTopN & ".csv"
	stProcAvgMemKTop = stBaseDir & stDateCompName & "-ProcAverage-MemK-Top" & nTopN & ".csv"
	stProcAvgPageFaultsTop = stBaseDir & stDateCompName & "-ProcAverage-PF-Top" & nTopN & ".csv"
	Set tsProcAvgCPUTop = oFso.OpenTextFile(stProcAvgCPUTop, ForWriting, True)
	Set tsProcAvgMemKTop = oFso.OpenTextFile(stProcAvgMemKTop, ForWriting, True)
	Set tsProcAvgPageFaultsTop = oFso.OpenTextFile(stProcAvgPageFaultsTop, ForWriting, True)
	tsProcAvgCPUTop.WriteLine stProcCPUout
	tsProcAvgPageFaultsTop.WriteLine stProcPageFaultsOut
	tsProcAvgMemKTop.WriteLine stProcMemKout
	tsProcAvgCPUTop.Close
	tsProcAvgPageFaultsTop.Close
	tsProcAvgMemKTop.Close
	Set tsProcAvgCPUTop = Nothing
	Set tsProcAvgPageFaultsTop = Nothing
	Set tsProcAvgMemKTop = Nothing
	
	arProcFiles = Array(stProcAvgCPUTop, stProcAvgMemKTop, stProcAvgPageFaultsTop, stCPU, stMemory, stNetwork)
	arProcSizes = Array(CInt(nTopN), CInt(nTopN) * 2, CInt(nTopN), _
			UBound(arCPUHeader), UBound(arMemHeader), UBound(arNetworkHeader))
	Wscript.Echo "Creating Excel File...."
	Wscript.Echo "Check behind Windows for Excel Compatibility Window"
	Set oWB = oExcel.Workbooks.Add
	For iXL = 0 To UBound(arProcFiles)
		stSheetName = ""
		nEnd = InStrRev(arProcFiles(iXL), ".")
		nDetails = InStr(1, arProcFiles(iXL), "ProcAverage", vbTextcompare)
		If nDetails Then
			nStart = InStr(nDetails, arProcFiles(iXL), "-", vbBinarycompare) + 1
			stSheetName = "Per Process " & Replace(Mid(arProcFiles(iXL), nStart, nEnd - nStart), "-", " ", 1, -1)
		Else
			nStart = InStrRev(arProcFiles(iXL), "-") + 1
			stSheetName = "System Wide " & Mid(arProcFiles(iXL), nStart, nEnd - nStart)
		End If
		Call AddWorkSheet(oExcel, arProcFiles(iXL), stSheetName)
		'Wscript.Echo "Processing range: " & stSheetName & vbTab & "params: " & CInt(iAvgTimeCount) + 1 & vbTab & arProcSizes(iXL) + 1
		Set oRange = oWB.ActiveSheet.Range("A1").Resize(CInt(iAvgTimeCount) + 1, arProcSizes(iXL) + 1)
		Call CreateChart(oExcel, stSheetName, oRange)
		Set oRange = Nothing
	Next
	Call AddWorkSheet(oExcel, stProcess, "Processes Reference")
	oWB.Sheets("Sheet1").Delete
	For Each oChart In oWB.Charts
		oChart.Move oWB.Sheets(1)
	Next
	stOutExcelFileName = ""
	stOutExcelFileName = stBaseDir & stDateCompName & ".xls"	
	oWB.SaveAs stOutExcelFileName, xlWorkbookNormal
	oWB.Close
	For Each stFile In arProcFiles
		oFso.DeleteFile(stFile)
	Next
	oFso.DeleteFile stProcess
	Erase arProcFiles
	Erase arProcSizes
	Redim arProcCPUAvg(1), arProcCPUtimeAvg(1), arProcMemKAvg(1), arProcPageFaultsAvg(1)
	Redim arProcCPUMax(1), arProcCPUtimeMax(1), arProcMemKMax(1), arProcPageFaultsMax(1)
	Redim arProcCPU(0), arProcCPUtime(0), arProcMemK(0), arProcPageFaults(0)
	Redim arCPU(3)
	Redim arNetwork(0)
	Redim arMemory(0)
	Redim arNetworkHeader(0), arMemHeader(0)
	Erase arCPUHeader
	Redim arCPUavg(0), arNetworkAvg(0), arMemoryAvg(0)
	dicProcHeader.RemoveAll
Next
oExcel.Quit
Set oExcel = Nothing

'******************************************************************************
Function FillMultiLine(byRef stLineIn, Byval stLineFill)
	If stLineIn = "" Then
		FillMultiLine = stLineFill
	Else
		FillMultiLine = stLineIn & vbCrlf & stLineFill
	End If
End Function
'******************************************************************************
'******************************************************************************
Sub CopyArray(Byval arIn, ByRef arOut, Byval nStart)
	For i = nStart To UBound(arIn)
		arOut(i) = arIn(i)
	Next
End Sub
'******************************************************************************
'******************************************************************************
Sub CopyAvgArray(Byval arIn, ByRef arOut, Byval nStart)
	For k = nStart To UBound(arIn)
		If IsNumeric(arIn(k)) Then
			arOut(k) = (arOut(k) + CSng(arIn(k)))' / 2.0
		End If
	Next
End Sub
'******************************************************************************
'******************************************************************************
Sub ArrayMax(Byval arIn, ByRef arMax, Byval nStart)
	For l = nStart To UBound(arIn)
	'	If arMax(l) = "" Then arMax(l) = arIn(l)
		If InStr(1, arIn(l), "~", vbBinarycompare) Then
			arTmp = Split(arIn(l), "~")
			arTmpMax = Split(arMax(l), "~")
			If UBound(arTmpMax) = UBound(arTmp) Then
				For a = 0 To UBound(arTmp)
					If CSng(arTmp(a)) > CSng(arTmpMax(a)) Then arTmpMax(a) = arTmp(a)
				Next
				arMax(l) = Join(arTmpMax, "~")
			Else
				arMax(l) = arIn(l)
			End If
		Else
			If IsNumeric(arIn(l)) Then
				If CSng(arIn(l)) > CSng(arMax(l)) Then arMax(l) = arIn(l)
			End If
		End If
	Next
End Sub
'******************************************************************************
'************************************************************************************
Sub SortArrayInd(ByRef arSort, Byval nStart, Byval bSortOrder, Byval nTopN)
	'bSortOrder - if true sort asc, otherwise desc
	Dim arTemp()
	If CInt(nTopN) > CInt(UBound(arSort)) Then nTopN = CInt(UBound(arSort))
	If InStr(1, arSort(nStart), "~", vbBinarycompare) Then
		bSpecial = True
		' Special case -- memory array holds two values in each element,
		' separated by "~" character
		' Separating values and placing into own array element for sort to work
		stTemp = Join(arSort, "~")
		arTmpS = Split(stTemp, "~")
		nMax = UBound(arTmpS)
		Redim arSort(nMax)
		For iTmp = 0 To nMax
			arSort(iTmp) = arTmpS(iTmp)
			'Wscript.Echo "arSort(" & iTmp & ") = " & arSort(iTmp)
		Next
		Erase arTmpS
		Redim arTemp(nMax)
		' Indexes should be remained same to be able to separate original array
		arTemp(0) = 0
		For i = (nStart * 2) To nMax Step 2
			arTemp(i) = Int(i / 2)
			arTemp(i - 1) = Int(i / 2)
		Next
	Else
		bSpecial = False
		nMax = UBound(arSort)
		Redim arTemp(nMax)
		For i = 0 To nMax
			arTemp(i) = i
		Next
	End If
	For i = nStart To nMax
		If isNumeric(arSort(i)) Then
			best_value = CSng(arSort(i))
		'	If Not isNumeric(arSort(i)) Then Wscript.Echo "i arSort(" & i & ") = " & arSort(i)
			best_j = i
	       	For j = i + 1 To nMax
	       		If IsNumeric(arSort(j)) Then
		       		If bSortOrder Then
						If CSng(arSort(j)) < CSng(best_value) Then
							best_value = CSng(arSort(j))
							best_j = j
						End If
					Else
						If CSng(arSort(j)) > CSng(best_value) Then
							best_value = CSng(arSort(j))
							best_j = j
						End If
					End If
				Else
		       		If bSortOrder Then
						If CSng(0) < CSng(best_value) Then
							best_value = CSng(0)
							best_j = j
						End If
					Else
						If CSng(0) > CSng(best_value) Then
							best_value = CSng(0)
							best_j = j
						End If
					End If
				End If
			Next
			best_i = CInt(arTemp(i))
			arTemp(i) = CInt(arTemp(best_j))
			arTemp(best_j) = CInt(best_i)
			arSort(best_j) = CSng(arSort(i))
			arSort(i) = CSng(best_value)
		Else
			best_value = 0
		End If
	Next
	Redim arSort(CInt(nTopN))
	If bSpecial Then
		iDist = 0
		iOrig = 0
		Do While iDist < CInt(nTopN + 1)
			If Ubound(Filter(arSort, arTemp(iOrig), True, vbTextcompare)) < 0 Then
				arSort(iDist) = arTemp(iOrig)
				iDist = iDist + 1
			End If
			iOrig = iOrig + 1
		Loop
	Else
		For i = 0 To nTopN
			arSort(i) = arTemp(i)
		Next
	End If
End Sub
'************************************************************************************
'************************************************************************************
Sub GetTopN(byRef stIn, Byval arMaxIn)
	arIn = Split(stIn, vbCrlf, -1, vbBinarycompare)
	Dim arTemp()
	Redim arTemp(UBound(arMaxIn))
	For iPr = 0 To UBound(arIn)
		arRow = Split(arIn(iPr), ",", -1, vbBinarycompare)
		For iM = 0 To Ubound(arMaxIn)
			If CInt(arMaxIn(iM)) > CInt(Ubound(arRow)) Then
				arTemp(iM) = 0
			Else
				arTemp(iM) = arRow(arMaxIn(iM))
			End If
		Next
		arIn(iPr) = Replace(Join(arTemp, ","), "~", ",", 1, -1, vbBinarycompare)
	Next
	stIn = Join(arIn, vbCrlf)
End Sub
'************************************************************************************
'************************************************************************************
Sub ProcDetails(Byval stLineProc)
	arLineProc = Split(stLineProc, ",", -1, vbBinarycompare)
	If bFirstProc Then
		dtStartProc = arLineProc(0)
		'Wscript.Echo "Start time set to: " & dtStartProc
		arProcCPUAvg(0) = arLineProc(0)
		arProcMemKAvg(0) = arLineProc(0)
		arProcPageFaultsAvg(0) = arLineProc(0)
		arProcCPUAvg(1) = arLineProc(3)
		arProcMemKAvg(1) = arLineProc(5) & "~" & arLineProc(7)
		arProcPageFaultsAvg(1) = arLineProc(6)
		arProcCPU(0) = arLineProc(0)
		arProcMemK(0) = arLineProc(0)
		arProcPageFaults(0) = arLineProc(0)
	'	arProcCPUtime(0) = arLineProc(0)
		bFirstProc = False
	End If
	If Not dicProcHeader.Exists(arLineProc(1) & ":" & arLineProc(2)) Then
		dicProcHeader.Add arLineProc(1) & ":" & arLineProc(2), iProc
		tsProcess.WriteLine arLineProc(1) & ":" & arLineProc(2) & "," & arLineProc(10)
		'Wscript.Echo "New header: " & arLineProc(1) & ":" & arLineProc(2) & ":" & arLineProc(10)
		iProc = iProc + 1
		If iProc > UBound(arProcCPU) Then
			Redim Preserve arProcCPU(iProc - 1)
	'		Redim Preserve arProcCPUtime(iProc)
			Redim Preserve arProcMemK(iProc - 1)
			Redim Preserve arProcPageFaults(iProc - 1)
		End If
	End If
	If DateDiff("s", arLineProc(0), arProcCPU(0)) = 0 Then
		arProcCPU(0) = arLineProc(0)
	'	arProcCPUtime(0) = arLineProc(0)
		arProcMemK(0) = arLineProc(0)
		arProcPageFaults(0) = arLineProc(0)
		arProcCPU(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(3)
	'	arProcCPUtime(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(4)
		arProcMemK(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(5) & "~" & arLineProc(7)
		arProcPageFaults(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(6)
	Else
		'New time entry -- need to flush data to file or to avg array
		arProcCPU(0) = arLineProc(0)
		arProcMemK(0) = arLineProc(0)
		arProcPageFaults(0) = arLineProc(0)
		Redim Preserve arProcCPUAvg(UBound(arProcCPU))
		Redim Preserve arProcMemKAvg(UBound(arProcMemK))
		Redim Preserve arProcPageFaultsAvg(UBound(arProcPageFaults))
		Redim Preserve arProcCPUMax(UBound(arProcCPU))
		Redim Preserve arProcMemKMax(UBound(arProcMemK))
		Redim Preserve arProcPageFaultsMax(UBound(arProcPageFaults))
	'	For averaging checking whether current entry should be included into
	'	averaged array or should become a first record in the next averaging
	'	iteration:
		If bAveraged Then
		'	dtStartProc = arLineProc(0)
		'	Calculating Max for each entry in array for future "top N" report
			arProcCPUMax(0) = 0
			arProcMemKMax(0) = 0
			arProcPageFaultsMax(0) = 0
			Call DivideArray(arProcCPUAvg, nAvgNumber, 1)
			Call DivideArray(arProcMemKAvg, nAvgNumber, 1)
			Call DivideArray(arProcPageFaultsAvg, nAvgNumber, 1)
			Call ArrayMax(arProcCPUAvg, arProcCPUMax, 1)
			Call ArrayMax(arProcMemKAvg, arProcMemKMax, 1)
			Call ArrayMax(arProcPageFaultsAvg, arProcPageFaultsMax, 1)
		'	Dumping result to output string (multiline) instead of file
		'	to be able to flush it to file after header is completed
			stProcCPUout = FillMultiLine(stProcCPUout, Join(arProcCPUAvg, ","))
			stProcMemKout = FillMultiLine(stProcMemKout, Join(arProcMemKAvg, ","))
			stProcPageFaultsOut = FillMultiLine(stProcPageFaultsOut, Join(arProcPageFaultsAvg, ","))
		'	Re-initializing average arrays to current value for next interval
			Call CopyArray(arProcCPU, arProcCPUAvg, 1)
		'	Call CopyArray(arProcCPUtime, arProcCPUtimeAvg, 1)
			Call CopyArray(arProcMemK, arProcMemKAvg, 1)
			Call CopyArray(arProcPageFaults, arProcPageFaultsAvg, 1)
		Else
			arProcCPUAvg(0) = arProcCPU(0)
			arProcMemKAvg(0) = arProcMemK(0)
			arProcPageFaultsAvg(0) = arProcPageFaults(0)
			Call CopyAvgArray(arProcCPU, arProcCPUAvg, 1)
		'	The following doesn't work for datetime variables:	
		'	Call CopyAvgArray(arProcCPUtime, arProcCPUtimeAvg, 1)
			For i = 1 To UBound(arProcMemK)
				arTemp = Split(arProcMemK(i), "~")
				If arProcMemKAvg(i) = "" Then arProcMemKAvg(i) = "0~0"
				arTempAvg = Split(arProcMemKAvg(i), "~")
				Call CopyAvgArray(arTemp, arTempAvg, 0)
				arProcMemKAvg(i) = Join(arTempAvg, "~")
			Next
			Call CopyAvgArray(arProcPageFaults, arProcPageFaultsAvg, 1)
			Call ZeroArray(arProcCPU, 1)
			Call ZeroArray(arProcMemK, 1)
			Call ZeroArray(arProcPageFaults, 1)
			'Wscript.Echo "arLine type: " & TypeName(arLine) & vbTab & arLine
			arProcCPU(0) = arLineProc(0)
		'	arProcCPUtime(0) = arLineProc(0)
			arProcMemK(0) = arLineProc(0)
			arProcPageFaults(0) = arLineProc(0)
			arProcCPU(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(3)
		'	arProcCPUtime(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(4)
			arProcMemK(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(5) & "~" & arLineProc(7)
			arProcPageFaults(dicProcHeader(arLineProc(1) & ":" & arLineProc(2))) = arLineProc(6)
		End If
	End If
End Sub
'************************************************************************************
'***********************************************************************************
Sub AddWorkSheet(ByRef oXL, Byval stFilePath, Byval stSName)
	Set oNewWB = oExcel.Workbooks.Open(stFilePath)
	oNewWB.ActiveSheet.Name = stSName
	oXL.Sheets(stSName).Move oWB.Sheets(1)
	Set oNewWB = Nothing
End Sub
'***********************************************************************************
'***********************************************************************************
Sub CreateChart(ByRef oXL, Byval stSName, ByRef oRng)
	With oXL
		.Charts.Add
		.ActiveChart.ChartType = xlLine
		.ActiveChart.SetSourceData oRng, xlColumns
		.ActiveChart.Location xlLocationAsNewSheet, stSName & " Chart"
	End With
	Set oChart = oXL.Charts(stSName & " Chart")
	With oChart
		.HasTitle = True
		.ChartTitle.Characters.Text = stSName
		.Axes(xlCategory, xlPrimary).HasTitle = False
		.Axes(xlValue, xlPrimary).HasTitle = False
		.HasAxis(xlCategory, xlPrimary) = True
		.HasAxis(xlValue, xlPrimary) = True
		.Axes(xlCategory, xlPrimary).CategoryType = xlCategoryScale
	End With
	oChart.PlotArea.Select
	With oChart.PlotArea
		.Border.Weight = xlThin
		.Border.LineStyle = xlNone
		.Interior.ColorIndex = xlNone
	End With
	oChart.Deselect
End Sub
'***********************************************************************************
'************************************************************************************
Sub ZeroArray(byRef arIn, Byval nStart)
	For iZero = nStart To UBound(arIn)
		If IsNumeric(arIn(iZero)) Then
			arIn(iZero) = 0
		Else
			arIn(iZero) = ""
		End If
	Next
End Sub
'************************************************************************************
'************************************************************************************
Sub DivideArray(byRef arIn, Byval nTimes, Byval nStart)
	For iDiv = nStart To UBound(arIn)
		If InStr(1, arIn(iDiv), "~", vbBinarycompare) Then
			arTmp = Split(arIn(iDiv), "~")
			For a = 0 To UBound(arTmp)
				arTmp(a) = CSng(arTmp(a)) / nTimes
			Next
			arIn(iDiv) = Join(arTmp, "~")
		Else
			If IsNumeric(arIn(iDiv)) Then
				arIn(iDiv) = CSng(arIn(iDiv)) / nTimes
			End If
		End If
	Next
End Sub
'************************************************************************************

stFinish = Timer
Wscript.Echo "Script ran: " & CSng(stFinish) - CSng(stStart) & " seconds. Number of entries processed: " & iTimeCount + 1 & _
		" Number of output samples: " & iAvgTimeCount
