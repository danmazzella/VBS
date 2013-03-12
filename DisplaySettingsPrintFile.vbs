On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Set fso = CreateObject("Scripting.FileSystemObject")

Set oArgs = Wscript.Arguments
If oArgs.Count = 1 Then
        strComputer = CStr(oArgs(0))
Else
    strComputer = inputBox("Enter computer name")
End If



'arrComputers = Array("JC1WDTRA2907")
'For Each strComputer In arrComputers
WScript.Echo
WScript.Echo "=========================================="
WScript.Echo "Computer: " & strComputer
WScript.Echo "=========================================="

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController", "WQL", _
									  wbemFlagReturnImmediately + wbemFlagForwardOnly)

For Each objItem In colItems
	count = count + 1
	Set WriteStuff = FSO.OpenTextFile("c:\users\dmazzell\desktop\monitor.txt", 8, True)
	
	if count = 1 then
		WriteStuff.WriteLine("Monitor 1")
	elseif counter = 2 then
		WriteStuff.WriteLine("Monitor 2")
	elseif counter = 3 then
		WriteStuff.WriteLine("Monitor 3")
	elseif counter = 4 then
		WriteStuff.WriteLine("Monitor 4")
	end if
	
	strAcceleratorCapabilities = Join(objItem.AcceleratorCapabilities, ",")
	WriteStuff.WriteLine("AcceleratorCapabilities: " & strAcceleratorCapabilities)
	WriteStuff.WriteLine("AdapterCompatibility: " & objItem.AdapterCompatibility)
	WriteStuff.WriteLine("AdapterDACType: " & objItem.AdapterDACType)
	WriteStuff.WriteLine("AdapterRAM: " & objItem.AdapterRAM)
	WriteStuff.WriteLine("Availability: " & objItem.Availability)
	strCapabilityDescriptions = Join(objItem.CapabilityDescriptions, ",")
	WriteStuff.WriteLine("CapabilityDescriptions: " & strCapabilityDescriptions)
	WriteStuff.WriteLine("Caption: " & objItem.Caption)
	WriteStuff.WriteLine("ColorTableEntries: " & objItem.ColorTableEntries)
	WriteStuff.WriteLine("ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode)
	WriteStuff.WriteLine("ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig)
	WriteStuff.WriteLine("CreationClassName: " & objItem.CreationClassName)
	WriteStuff.WriteLine("CurrentBitsPerPixel: " & objItem.CurrentBitsPerPixel)
	WriteStuff.WriteLine("CurrentHorizontalResolution: " & objItem.CurrentHorizontalResolution)
	WriteStuff.WriteLine("CurrentNumberOfColors: " & objItem.CurrentNumberOfColors)
	WriteStuff.WriteLine("CurrentNumberOfColumns: " & objItem.CurrentNumberOfColumns)
	WriteStuff.WriteLine("CurrentNumberOfRows: " & objItem.CurrentNumberOfRows)
	WriteStuff.WriteLine("CurrentRefreshRate: " & objItem.CurrentRefreshRate)
	WriteStuff.WriteLine("CurrentScanMode: " & objItem.CurrentScanMode)
	WriteStuff.WriteLine("CurrentVerticalResolution: " & objItem.CurrentVerticalResolution)
	WriteStuff.WriteLine("Description: " & objItem.Description)
	WriteStuff.WriteLine("DeviceID: " & objItem.DeviceID)
	WriteStuff.WriteLine("DeviceSpecificPens: " & objItem.DeviceSpecificPens)
	WriteStuff.WriteLine("DitherType: " & objItem.DitherType)
	WriteStuff.WriteLine("DriverDate: " & WMIDateStringToDate(objItem.DriverDate))
	WriteStuff.WriteLine("DriverVersion: " & objItem.DriverVersion)
	WriteStuff.WriteLine("ErrorCleared: " & objItem.ErrorCleared)
	WriteStuff.WriteLine("ErrorDescription: " & objItem.ErrorDescription)
	WriteStuff.WriteLine("ICMIntent: " & objItem.ICMIntent)
	WriteStuff.WriteLine("ICMMethod: " & objItem.ICMMethod)
	WriteStuff.WriteLine("InfFilename: " & objItem.InfFilename)
	WriteStuff.WriteLine("InfSection: " & objItem.InfSection)
	WriteStuff.WriteLine("InstallDate: " & WMIDateStringToDate(objItem.InstallDate))
	WriteStuff.WriteLine("InstalledDisplayDrivers: " & objItem.InstalledDisplayDrivers)
	WriteStuff.WriteLine("LastErrorCode: " & objItem.LastErrorCode)
	WriteStuff.WriteLine("MaxMemorySupported: " & objItem.MaxMemorySupported)
	WriteStuff.WriteLine("MaxNumberControlled: " & objItem.MaxNumberControlled)
	WriteStuff.WriteLine("MaxRefreshRate: " & objItem.MaxRefreshRate)
	WriteStuff.WriteLine("MinRefreshRate: " & objItem.MinRefreshRate)
	WriteStuff.WriteLine("Monochrome: " & objItem.Monochrome)
	WriteStuff.WriteLine("Name: " & objItem.Name)
	WriteStuff.WriteLine("NumberOfColorPlanes: " & objItem.NumberOfColorPlanes)
	WriteStuff.WriteLine("NumberOfVideoPages: " & objItem.NumberOfVideoPages)
	WriteStuff.WriteLine("PNPDeviceID: " & objItem.PNPDeviceID)
	strPowerManagementCapabilities = Join(objItem.PowerManagementCapabilities, ",")
	WriteStuff.WriteLine("PowerManagementCapabilities: " & strPowerManagementCapabilities)
	WriteStuff.WriteLine("PowerManagementSupported: " & objItem.PowerManagementSupported)
	WriteStuff.WriteLine("ProtocolSupported: " & objItem.ProtocolSupported)
	WriteStuff.WriteLine("ReservedSystemPaletteEntries: " & objItem.ReservedSystemPaletteEntries)
	WriteStuff.WriteLine("SpecificationVersion: " & objItem.SpecificationVersion)
	WriteStuff.WriteLine("Status: " & objItem.Status)
	WriteStuff.WriteLine("StatusInfo: " & objItem.StatusInfo)
	WriteStuff.WriteLine("SystemCreationClassName: " & objItem.SystemCreationClassName)
	WriteStuff.WriteLine("SystemName: " & objItem.SystemName)
	WriteStuff.WriteLine("SystemPaletteEntries: " & objItem.SystemPaletteEntries)
	WriteStuff.WriteLine("TimeOfLastReset: " & WMIDateStringToDate(objItem.TimeOfLastReset))
	WriteStuff.WriteLine("VideoArchitecture: " & objItem.VideoArchitecture)
	WriteStuff.WriteLine("VideoMemoryType: " & objItem.VideoMemoryType)
	WriteStuff.WriteLine("VideoMode: " & objItem.VideoMode)
	WriteStuff.WriteLine("VideoModeDescription: " & objItem.VideoModeDescription)
	WriteStuff.WriteLine("VideoProcessor: " & objItem.VideoProcessor)
  
	WriteStuff.WriteLine("End of monitor " & count)
	WriteStuff.WriteLine("")
	WriteStuff.WriteLine("")
	SET WriteStuff = NOTHING
  
Next

Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Print Monitor Settings To File"""
msgbox("Script Complete")

Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function
