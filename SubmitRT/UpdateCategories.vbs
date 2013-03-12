TheGroups = Array("linux","windows","monitoring","security","network","datacenter","dba","voice","londonit","pcteam")
For Counting = 0 to 9
	'Set your settings
	strFileURL = "http://tickets/forms/csvgroups.php?group=" & TheGroups(counting)
	strHDLocation = "\\jc1wsalt03\library\Packages\Dantools\SubmitRT\" & TheGroups(counting) & ".csv"
	
	' Fetch the file

	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")

	objXMLHTTP.open "GET", strFileURL, false
	objXMLHTTP.send()

	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1 'adTypeBinary

		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0    'Set the stream position to the start

		Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
		Set objFSO = Nothing

		objADOStream.SaveToFile strHDLocation
		objADOStream.Close
		Set objADOStream = Nothing
	End if

	Set objXMLHTTP = Nothing
Next