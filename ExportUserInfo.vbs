' XLWrite.vbs
' VBScript program demonstrating how to write values to cells in a
' Microsoft Excel spreadsheet.
'
' ----------------------------------------------------------------------
' Copyright (c) 2002-2010 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - October 12, 2002
' Version 1.1 - February 19, 2003 - Standardize Hungarian notation.
' Version 1.2 - January 25, 2004 - Modify error trapping.
' Version 1.3 - April 29, 2010 - Specify FileFormat of spreadsheet.
' Version 1.4 - November 6, 2010 - No need to set objects to Nothing.
' This program documents a few user object attributes, including the
' group memberships, to an Excel spreadsheet. If the spreadsheet file
' already exists, the user is asked by Excel if they want to replace the
' file.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objUser, strExcelPath, objExcel, objSheet, k, objGroup

Const xlExcel7 = 39

' User object whose group membership will be documented in the
' spreadsheet.
Set objUser = GetObject("LDAP://cn=Daniel Mazzella,ou=Backoffice,ou=User Accounts,dc=global,dc=knight,dc=com")

' Spreadsheet file to be created.
strExcelPath = "c:\temp\UserGroup.xls"

' Bind to Excel object.
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If (Err.Number <> 0) Then
    On Error GoTo 0
    Wscript.Echo "Excel application not found."
    Wscript.Quit
End If
On Error GoTo 0

' Create a new workbook.
objExcel.Workbooks.Add

' Bind to worksheet.
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.Name = "User Groups"

' Populate spreadsheet cells with user attributes.
objSheet.Cells(1, 1).Value = "User Common Name"
objSheet.Cells(2, 1).Value = "sAMAccountName"
objSheet.Cells(3, 1).Value = "Display Name"
objSheet.Cells(4, 1).Value = "Distinguished Name"
objSheet.Cells(1, 2).Value = objUser.cn
objSheet.Cells(2, 2).Value = objUser.sAMAccountName
objSheet.Cells(3, 2).Value = objUser.displayName
objSheet.Cells(4, 2).Value = objUser.distinguishedName
objSheet.Cells(5, 1).Value = "Groups"

' Enumerate groups and add group names to spreadsheet.
k = 5
For Each objGroup In objUser.Groups
    objSheet.Cells(k, 2).Value = objGroup.sAMAccountName
    k = k + 1
Next

' Format the spreadsheet.
objSheet.Range("A1:A5").Font.Bold = True
objSheet.Select
objSheet.Range("B5").Select
objExcel.ActiveWindow.FreezePanes = True
objExcel.Columns(1).ColumnWidth = 20
objExcel.Columns(2).ColumnWidth = 30

' Save the spreadsheet and close the workbook.
' Specify Excel7 File Format.
objExcel.ActiveWorkbook.SaveAs strExcelPath, xlExcel7
objExcel.ActiveWorkbook.Close

' Quit Excel.
objExcel.Application.Quit

Wscript.Echo "Done"
