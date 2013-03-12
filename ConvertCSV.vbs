Option Explicit
Const vbNormal = 1

DIM objXL, objWb, objR     ' Excel object variables
DIM Title, Text, tmp, i, j, file, name, savename

file = "c:\temp\LaptopLoaner.csv"
name = "LaptopLoaner"

savename = "c:\temp\LaptopLoaner.xls"

Function GetPath
' Retrieve the script path
 DIM path
 path = WScript.ScriptFullName  ' Script name
 GetPath = Left(path, InstrRev(path, "\"))
End Function



' create an Excel object reference
Set objXL = WScript.CreateObject ("Excel.Application")

objXL.WindowState = vbNormal ' Normal
objXL.Height = 300           ' height
objXL.Width = 400            ' width
objXL.Left = 40              ' X-Position
objXL.Top = 20               ' Y-Position
objXL.Visible = false		' show window


' Load the Excel file from the script's folder
Set objWb = objXl.WorkBooks.Open(file)

' Get the loaded worksheet object
Set objWb = objXL.ActiveWorkBook.WorkSheets("LaptopLoaner")
objWb.Activate               ' not absolutely necessary (for CSV)

'WScript.Echo "worksheet imported"


' turn of those annoying warning messages
OBJXL.DISPLAYALERTS = fALSE

'wscript.echo savename

' xlWorkbookNormal
objxl.ActiveWorkbook.SaveAs savename, &HFFFFEFD1


objXl.Quit()

Set objXL = Nothing
Set objWB = Nothing
Set objR = Nothing
