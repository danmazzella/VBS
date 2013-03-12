Dim x, zz 
Set objRoot = GetObject("LDAP://RootDSE") 
strDNC = objRoot.Get("DefaultNamingContext") 
Set objDomain = GetObject("LDAP://" & strDNC) ' Bind to the top of the Domain using LDAP using ROotDSE 
MsgBox "You will receive a popup when complete, The file will be located at C:\Temp\AD.csv"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
If FileSystem.FileExists("c:\temp\AD.csv") then
	FileSystem.DeleteFile("c:\temp\AD.csv")
End If
TheColumns = "SamAccountName,CN,FirstName,LastName,Initials,Descrip,Office,Telephone,Email,WebPage,Addr1,City,State,ZipCode,Title,Department,Company,Manager,Profile,LoginScript,HomeDirectory,HomeDrive,Adspath,PrimarySMTP"
Set WriteStuff = FileSystem.OpenTextFile("c:\temp\AD.csv", 8, True)
WriteStuff.WriteLine(TheColumns)
WriteStuff.Close
SET WriteStuff = NOTHING
SET FileSystem = NOTHING  

Call enummembers(objDomain) 

Sub enumMembers(objDomain) 
	On Error Resume Next 
	Dim Secondary(20) ' Variable to store the Array of 2ndary email alias's 
	For Each objMember In objDomain ' go through the collection 
		If ObjMember.Class = "user" Then ' if not User object, move on. 
			x = x +1 ' counter used to increment the cells in Excel 

			' I set AD properties to variables so if needed you could do Null checks or add if/then's to this code 
			' this was done so the script could be modified easier. 
			SamAccountName = ObjMember.samAccountName 
			Cn = ObjMember.CN 
			FirstName = objMember.GivenName 
			LastName = objMember.sn 
			initials = objMember.initials 
			Descrip = objMember.description 
			Office = objMember.physicalDeliveryOfficeName 
			Telephone = objMember.telephonenumber 
			EmailAddr = objMember.mail 
			WebPage = objMember.wwwHomePage 
			Addr1 = objMember.streetAddress 
			City = objMember.l 
			State = objMember.st 
			ZipCode = objMember.postalCode 
			Title = ObjMember.Title 
			Department = objMember.Department 
			Company = objMember.Company 
			Manager = ObjMember.Manager 
			Profile = objMember.profilePath 
			LoginScript = objMember.scriptpath 
			HomeDirectory = ObjMember.HomeDirectory 
			HomeDrive = ObjMember.homeDrive 
			AdsPath = Objmember.Adspath 

			AllAttributes = Array(SamAccountName,Cn,FirstName,LastName,initials,Descrip,Office,Telephone,EmailAddr,WebPage,Addr1,City,State,ZipCode,Title,Department,Company,Manager,Profile,LoginScript,HomeDirectory,HomeDrive,AdsPath)
			
			zz = 1 ' Counter for array of 2ndary email addresses 
			For each email in ObjMember.proxyAddresses 
				If Left (email,5) = "SMTP:" Then 
					Primary = Mid (email,6) ' if SMTP is all caps, then it's the Primary 
				ElseIf Left (email,5) = "smtp:" Then 
					Secondary(zz) = Mid (email,6) ' load the list of 2ndary SMTP emails into Array. 
					zz = zz + 1 
				End If 
			Next 
			' Write the values to Excel, using the X counter to increment the rows. 

			' Write out the Array for the 2ndary email addresses. 
			Counter = 0
			For ll = 1 To 20 
				if Secondary(ll) <> "" then
					if Counter = 0 then
						SecondEmail = Secondary(ll)
					else
						SecondEmail = SecondEmail + ", " + Secondary(ll)
					End if
					Counter = Counter + 1
				end if
			Next 
			SecondEmail = """" & SecondEmail & """"
			
			If SamAccountName <> " " AND EmailAddr <> "" then
				Set FileSystem = CreateObject("Scripting.FileSystemObject")
				Set WriteStuff = FileSystem.OpenTextFile("c:\temp\AD.csv", 8, True)
				Counter = 0
				for each x in AllAttributes
					if Counter = 0 then
						TheFile = TheFile + """" + x + """"
					Else
						TheFile = TheFile + "," + """" + x + """"
					End if
					Counter = Counter + 1
					x = x + 1
				Next
				TheFile = TheFile + "," + SecondEmail
				WriteStuff.WriteLine(TheFile)
				WriteStuff.Close
				TheFile = ""
				SET WriteStuff = NOTHING
				SET FileSystem = NOTHING    
			End if
			
			' Blank out Variables in case the next object doesn't have a value for the property 
			SamAccountName = "-" 
			Cn = "-" 
			FirstName = "-" 
			LastName = "-" 
			initials = "-" 
			Descrip = "-" 
			Office = "-" 
			Telephone = "-" 
			EmailAddr = "-" 
			WebPage = "-" 
			Addr1 = "-" 
			City = "-" 
			State = "-" 
			ZipCode = "-" 
			Title = "-" 
			Department = "-" 
			Company = "-" 
			Manager = "-" 
			Profile = "-" 
			LoginScript = "-" 
			HomeDirectory = "-" 
			HomeDrive = "-" 
			AdsPath = "-"
			
			For ll = 1 To 20 
				Secondary(ll) = "" 
			Next 
			Primary = "-" 
			SecondEmail = ""
		End If 

		' If the AD enumeration runs into an OU object, call the Sub again to itinerate 

		If objMember.Class = "organizationalUnit" or OBjMember.Class = "container" Then 
			enumMembers (objMember) 
		End If 
	Next 
End Sub 

Sub ExcelSetup(shtName) ' This sub creates an Excel worksheet and adds Column heads to the 1st row 
	Set objExcel = CreateObject("Excel.Application") 
	Set objwb = objExcel.Workbooks.Add 
	Set objwb = objExcel.ActiveWorkbook.Worksheets(shtName) 
	Objwb.Name = "Active Directory Users" ' name the sheet 
	objwb.Activate 
	objExcel.Visible = True 
	objwb.Cells(1, 2).Value = "SamAccountName" 
	objwb.Cells(1, 3).Value = "CN" 
	objwb.Cells(1, 4).Value = "FirstName" 
	objwb.Cells(1, 5).Value = "LastName" 
	objwb.Cells(1, 6).Value = "Initials" 
	objwb.Cells(1, 7).Value = "Descrip" 
	objwb.Cells(1, 8).Value = "Office" 
	objwb.Cells(1, 9).Value = "Telephone" 
	objwb.Cells(1, 10).Value = "Email" 
	objwb.Cells(1, 11).Value = "WebPage" 
	objwb.Cells(1, 12).Value = "Addr1" 
	objwb.Cells(1, 13).Value = "City" 
	objwb.Cells(1, 14).Value = "State" 
	objwb.Cells(1, 15).Value = "ZipCode" 
	objwb.Cells(1, 16).Value = "Title" 
	objwb.Cells(1, 17).Value = "Department" 
	objwb.Cells(1, 18).Value = "Company" 
	objwb.Cells(1, 19).Value = "Manager" 
	objwb.Cells(1, 20).Value = "Profile" 
	objwb.Cells(1, 21).Value = "LoginScript" 
	objwb.Cells(1, 22).Value = "HomeDirectory" 
	objwb.Cells(1, 23).Value = "HomeDrive" 
	objwb.Cells(1, 24).Value = "Adspath" 
	objwb.Cells(1, 25).Value = "LastLogin" 
	objwb.Cells(1, 26).Value = "Primary SMTP" 
	TheColumns = "SamAccountName,CN,FirstName,LastName,Initials,Descrip,Office,Telephone,Email,WebPage,Addr1,City,State,ZipCode,Title,Department,Company,Manager,Profile,LoginScript,HomeDirectory,HomeDrive,Adspath,LastLogin,PrimarySMTP"
End Sub 

Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Export All of Active Directory"""
MsgBox "Done" ' show that script is complete 