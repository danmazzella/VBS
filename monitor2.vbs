Dim MyArray() ' définition d'un tableau de variable qui aura comme éléments les noms de machines à traiter.
Dim path_txt, path_File_monitor, temp, FileToDayTxt
MsgBox "This program requires a file C:\Temp\ComputerNames.txt"
MsgBox "And exports data to c:\Temp\Monitor.csv"
path_txt = "C:\Temp\ComputerNames.txt" 
ReDim MyArray(GetNumberOfLines(path_txt)) ' array redimensionné en fct de lecteure nb lignes dans fichier ListePC.txt
Call Lis_Fichier(path_txt, 1) ' maintenant "array" est rempli
temp = Replace (Date, "/", "_") ' remplace / par - car non accepté ds les noms de fichier par Windows
path_File_monitor = "c:\temp\Monitors - " & temp & ".csv" ' ajoute la date au nom du fichier
Call EcritFichier (path_File_monitor, "Computer" & "," & "Model" & "," & "Serial #" & "," & "Vendor ID" & "," & "Manufacture Date" & "," & "Messages", 2, true) ' préparation du fichier

' ------------------------------------------------------------------------------------------------------------------------------------
For i = 0 To UBound(MyArray) - 1 ' debut de la boucle
On error resume next

'this code is based on the EEDID spec found at http://www.vesa.org
'and by my hacking around in the windows registry
'the code was tested on WINXP,WIN2K and WIN2K3
'it should work on WINME and WIN98SE
'It should work with multiple monitors, but that hasn't been tested either.
'*****************************************************************************************
'
'*****************************************************************************************
'It should be noted that this is not 100% reliable
'I have witnessed occasions where for one reason or another windows
'can't or doesn't read the EDID info at boot (example would be someone
'booting with the monitor turned off) and so windows changes the active
'monitor to "Default_Monitor"
'Another reason for reliability problems is that there is no
'requirement in the EDID spec that a manufacture include the
'serial number in the EDID data AND only EDIDv1.2 and beyond
'have a requirement that the EDID contain a descriptive
'model name
'That being said, here goes....
'*****************************************************************************************
'
'*****************************************************************************************
'Monitors are stored in HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\
'
'Unfortunately, not only monitors are stored here Video Chipsets and maybe some other stuff
'is also here.
'
'Monitors in "HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\" are organized like this:
' HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID>\<PNP_ID>\
'Since not only monitors will be found under DISPLAY sub key you need to find out which
'devices are monitors.
'This can be deterimined by looking at the value "HardwareID" located
'at HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\<PNP_ID>\
'If the device is a monitor then the "HardwareID" value will contain the data "Monitor\<VESA_Monitor_ID>"
'
'The Next difficulty is that all monitors are stored here not just the one curently plugged in.
'So, If you ever switched monitors the old one(s) will still be in the registry.
'You can tell which monitor(s) are active because they will have a sub-key named "Control"
'*****************************************************************************************
'
Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
Dim strComputer, message

Dim intMonitorCount
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3
Dim sValue
dim i, iRC, iRC2, iRC3
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintEDID
Dim strRawEDID
Dim ByteValue, strSerFind, strMdlFind
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmp, tmpser, tmpmdl, tmpctr

strComputer = MyArray(i)
If strcomputer = "" Then WScript.Quit
strComputer = UCase(strComputer)
wscript.echo strComputer

Dim strarrRawEDID()
intMonitorCount=0
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'get a handle to the WMI registry object
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "/root/default:StdRegProv")

If Err <> 0 Then
Call EcritFichier (path_File_monitor, strComputer & ",,,,," & "Failed", 8, false)
else


sBaseKey = "SYSTEM\CurrentControlSet\Enum\DISPLAY\"
'enumerate all the keys HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\
iRC = oRegistry.EnumKey(HKLM, sBaseKey, arSubKeys)
For Each sKey In arSubKeys
'we are now in the registry at the level of:
'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\
'we need to dive in one more level and check the data of the "HardwareID" value
sBaseKey2 = sBaseKey & sKey & "\"
iRC2 = oRegistry.EnumKey(HKLM, sBaseKey2, arSubKeys2)
For Each sKey2 In arSubKeys2
'now we are at the level of:
'HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\<VESA_Monitor_ID\<PNP_ID>\
'so we can check the "HardwareID" value
oRegistry.GetMultiStringValue HKLM, sBaseKey2 & sKey2 & "\", "HardwareID", sValue
for tmpctr=0 to ubound(svalue)
If lcase(left(svalue(tmpctr),8))="monitor\" then
'If it is a monitor we will check for the existance of a control subkey
'that way we know it is an active monitor
sBaseKey3 = sBaseKey2 & sKey2 & "\"
iRC3 = oRegistry.EnumKey(HKLM, sBaseKey3, arSubKeys3)
For Each sKey3 In arSubKeys3
'Kaplan edit
strRawEDID = ""
If skey3="Control" Then
'If the Control sub-key exists then we should read the edid info
oRegistry.GetBinaryValue HKLM, sbasekey3 & "Device Parameters\", "EDID", arrintEDID
If vartype(arrintedid) <> 8204 then 'and If we don't find it...
strRawEDID="EDID Not Available" 'store an "unavailable message
else
for each bytevalue in arrintedid 'otherwise conver the byte array from the registry into a string (for easier processing later)
strRawEDID=strRawEDID & chr(bytevalue)
Next
End If
'now take the string and store it in an array, that way we can support multiple monitors
redim preserve strarrRawEDID(intMonitorCount)
strarrRawEDID(intMonitorCount)=strRawEDID
intMonitorCount=intMonitorCount+1
End If
Next
End If
Next 
Next 
Next
'*****************************************************************************************
'now the EDID info for each active monitor is stored in an array of strings called strarrRawEDID
'so we can process it to get the good stuff out of it which we will store in a 5 dimensional array
'called arrMonitorInfo, the dimensions are as follows:
'0=VESA Mfg ID, 1=VESA Device ID, 2=MFG Date (M/YYYY),3=Serial Num (If available),4=Model Descriptor
'5=EDID Version
'*****************************************************************************************
On Error Resume Next
dim arrMonitorInfo()
redim arrMonitorInfo(intMonitorCount-1,5)
dim location(3)
for tmpctr=0 to intMonitorCount-1
If strarrRawEDID(tmpctr) <> "EDID Not Available" then
'*********************************************************************
'first get the model and serial numbers from the vesa descriptor
'blocks in the edid. the model number is required to be present
'according to the spec. (v1.2 and beyond)but serial number is not
'required. There are 4 descriptor blocks in edid at offset locations
'&H36 &H48 &H5a and &H6c each block is 18 bytes long
'*********************************************************************
location(0)=mid(strarrRawEDID(tmpctr),&H36+1,18)
location(1)=mid(strarrRawEDID(tmpctr),&H48+1,18)
location(2)=mid(strarrRawEDID(tmpctr),&H5a+1,18)
location(3)=mid(strarrRawEDID(tmpctr),&H6c+1,18)

'you can tell If the location contains a serial number If it starts with &H00 00 00 ff
strSerFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hff)
'or a model description If it starts with &H00 00 00 fc
strMdlFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hfc)

intSerFoundAt=-1
intMdlFoundAt=-1
for findit = 0 to 3
If instr(location(findit),strSerFind)>0 then
intSerFoundAt=findit
End If
If instr(location(findit),strMdlFind)>0 then
intMdlFoundAt=findit
End If
Next

'If a location containing a serial number block was found then store it
If intSerFoundAt<>-1 then
tmp=right(location(intSerFoundAt),14)
If instr(tmp,chr(&H0a))>0 then
tmpser=trim(left(tmp,instr(tmp,chr(&H0a))-1))
Else
tmpser=trim(tmp)
End If
'although it is not part of the edid spec it seems as though the
'serial number will frequently be preceeded by &H00, this
'compensates for that
If left(tmpser,1)=chr(0) then tmpser=right(tmpser,len(tmpser)-1)
else
tmpser="Not Found"
End If

'If a location containing a model number block was found then store it
If intMdlFoundAt<>-1 then
tmp=right(location(intMdlFoundAt),14)
If instr(tmp,chr(&H0a))>0 then
tmpmdl=trim(left(tmp,instr(tmp,chr(&H0a))-1))
else
tmpmdl=trim(tmp)
End If
'although it is not part of the edid spec it seems as though the
'serial number will frequently be preceeded by &H00, this
'compensates for that
If left(tmpmdl,1)=chr(0) then tmpmdl=right(tmpmdl,len(tmpmdl)-1)
else
tmpmdl="Not Found"
End If

'**************************************************************
'Next get the mfg date
'**************************************************************
Dim tmpmfgweek,tmpmfgyear,tmpmdt
'the week of manufacture is stored at EDID offset &H10
tmpmfgweek=asc(mid(strarrRawEDID(tmpctr),&H10+1,1))

'the year of manufacture is stored at EDID offset &H11
'and is the current year -1990
tmpmfgyear=(asc(mid(strarrRawEDID(tmpctr),&H11+1,1)))+1990

'store it in month/year format 
tmpmdt=month(dateadd("ww",tmpmfgweek,datevalue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear

'**************************************************************
'Next get the edid version
'**************************************************************
'the version is at EDID offset &H12
Dim tmpEDIDMajorVer, tmpEDIDRev, tmpVer
tmpEDIDMajorVer=asc(mid(strarrRawEDID(tmpctr),&H12+1,1))

'the revision level is at EDID offset &H13
tmpEDIDRev=asc(mid(strarrRawEDID(tmpctr),&H13+1,1))

'store it in month/year format 
tmpver=chr(48+tmpEDIDMajorVer) & "." & chr(48+tmpEDIDRev)

'**************************************************************
'Next get the mfg id
'**************************************************************
'the mfg id is 2 bytes starting at EDID offset &H08
'the id is three characters long. using 5 bits to represent
'each character. the bits are used so that 1=A 2=B etc..
'
'get the data
Dim tmpEDIDMfg, tmpMfg
Dim Char1, Char2, Char3
Dim Byte1, Byte2
tmpEDIDMfg=mid(strarrRawEDID(tmpctr),&H08+1,2) 
Char1=0 : Char2=0 : Char3=0
Byte1=asc(left(tmpEDIDMfg,1)) 'get the first half of the string
Byte2=asc(right(tmpEDIDMfg,1)) 'get the first half of the string
'now shift the bits
'shift the 64 bit to the 16 bit
If (Byte1 and 64) > 0 then Char1=Char1+16
'shift the 32 bit to the 8 bit
If (Byte1 and 32) > 0 then Char1=Char1+8
'etc....
If (Byte1 and 16) > 0 then Char1=Char1+4
If (Byte1 and 8) > 0 then Char1=Char1+2
If (Byte1 and 4) > 0 then Char1=Char1+1

'the 2nd character uses the 2 bit and the 1 bit of the 1st byte
If (Byte1 and 2) > 0 then Char2=Char2+16
If (Byte1 and 1) > 0 then Char2=Char2+8
'and the 128,64 and 32 bits of the 2nd byte
If (Byte2 and 128) > 0 then Char2=Char2+4
If (Byte2 and 64) > 0 then Char2=Char2+2
If (Byte2 and 32) > 0 then Char2=Char2+1

'the bits for the 3rd character don't need shifting
'we can use them as they are
Char3=Char3+(Byte2 and 16)
Char3=Char3+(Byte2 and 8)
Char3=Char3+(Byte2 and 4)
Char3=Char3+(Byte2 and 2)
Char3=Char3+(Byte2 and 1)
tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)

'**************************************************************
'Next get the device id
'**************************************************************
'the device id is 2bytes starting at EDID offset &H0a
'the bytes are in reverse order.
'this code is not text. it is just a 2 byte code assigned
'by the manufacturer. they should be unique to a model
Dim tmpEDIDDev1, tmpEDIDDev2, tmpDev

tmpEDIDDev1=hex(asc(mid(strarrRawEDID(tmpctr),&H0a+1,1)))
tmpEDIDDev2=hex(asc(mid(strarrRawEDID(tmpctr),&H0b+1,1)))
If len(tmpEDIDDev1)=1 then tmpEDIDDev1="0" & tmpEDIDDev1
If len(tmpEDIDDev2)=1 then tmpEDIDDev2="0" & tmpEDIDDev2
tmpdev=tmpEDIDDev2 & tmpEDIDDev1

'**************************************************************
'finally store all the values into the array
'**************************************************************
'Kaplan adds code to avoid duplication...

If Not InArray(tmpser,arrMonitorInfo,3) Then
arrMonitorInfo(tmpctr,0)=tmpmfg
arrMonitorInfo(tmpctr,1)=tmpdev
arrMonitorInfo(tmpctr,2)=tmpmdt
arrMonitorInfo(tmpctr,3)=tmpser
arrMonitorInfo(tmpctr,4)=tmpmdl
arrMonitorInfo(tmpctr,5)=tmpVer
End If
End If
Next

'For now just a simple screen print will suffice for output.
'But you could take this output and write it to a database or a file
'and in that way use it for asset management.
for tmpctr = 0 to intMonitorCount-1
If arrMonitorInfo(tmpctr,1) <> "" And arrMonitorInfo(tmpctr,0) <> "PNP" Then
Call EcritFichier (path_File_monitor, strComputer & "," & arrMonitorInfo(tmpctr,4) & "," & arrMonitorInfo(tmpctr,3)& "," & _
arrMonitorInfo(tmpctr,0) & "," & arrMonitorInfo(tmpctr,2), 8, false)
End If
Next
End If
Next
MyArray() = nothing
Msgbox ("End of script")

' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
Sub EcritFichier (strFile, strString, intMode, bool_create)
Set oFso = CreateObject("Scripting.FileSystemObject")
Set f = oFso.OpenTextFile(strFile, intMode, bool_create)
f.WriteLine(strString)
f.close
Set oFso = nothing
Set f = nothing
End Sub
' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
Function InArray(strValue,List,Col)
Dim i
For i = 0 to UBound(List)
If List(i,col) = cstr(strValue) Then
InArray = True
Exit Function
End If
Next
InArray = False
End Function
'------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
Function GetPath()
Dim path
path = WScript.ScriptFullName
GetPath = Left(path, InStrRev(path, "\"))
End Function
' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
Function GetNumberOfLines (strFile)
Dim oFso, f, i
Set oFso = CreateObject("Scripting.FileSystemObject")
Set f = oFso.OpenTextFile(strFile, 1)
i=0

Do 
f.ReadLine
i = i+1
Loop While Not f.AtEndOfStream 

f.Close
Set oFso = nothing
Set f = nothing
GetNumberOfLines = i
End Function
' ------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------
Sub Lis_Fichier(strfile, intmode)
Dim oFso, f, nb
Set oFso = CreateObject("Scripting.FileSystemObject")
Set f = oFso.OpenTextFile(strfile, intmode)
nb = GetNumberOfLines(strFile)

For i = 1 To nb
MyArray(i - 1) = f.ReadLine
Next 

f.Close
Set oFso = nothing
Set f = nothing
End Sub
' ------------------------------------------------------------------------------------------------------------------------------------


Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Check Monitors for List of PC"""
msgbox("Script Complete")