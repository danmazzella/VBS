'******************************************************************************
'Script to search for a user, users or groups in an Active Directory Domain.
'The script accepts wild cards to perform the search.
'
'The search is performed against the AD information stored in samAccountName
'
'Use this command line for help on script usage:
'
'cscript SearchDC.vbs /?
'
'SearchDC.vbs version 0.2 was last updated on 24/03/2003 by Dan Thomson
'E-mail address: - dethomson@hotmail.com
'
'This script was originally authored by Nigel Thomas on 28/06/2002
'E-mail address: - nthomas2020@yahoo.com
'
' All rights reserved.
'******************************************************************************

Option Explicit
On Error Resume Next

'Declare variables
Dim strSearch
Dim strAdsPath
Dim strServerName
Dim strDefaultDomainNC
Dim strADSQuery
Dim objQueryResultSet
Dim objADOConn
Dim objADOCommand
Dim objUser
Dim intCount


Const adStateOpen = 1

'Check for command line argument
If WScript.Arguments.Count < 1 Then
  Do While strSearch = ""
    'Prompt for search criteria
    strSearch = InputBox("Please specify the search criteria", "Search AD")
    'If no input...Check if user wants to exit
    'If user selects Yes then exit.
    If strSearch = "" Then If MsgBox("Do you wish to cancel ?", 36, _
      "Search AD") = 6 Then Call Cleanup(1)
  Loop
Else
  'Use search criteria specified on the command line
  strSearch = WScript.Arguments.Item(0)
End If

If strSearch = "/?" Then Call ShowUsage

'Get the local Computer Name
strServerName = WScript.CreateObject("WScript.Network").ComputerName

'Get the Default Domain Naming Context
strDefaultDomainNC = GetObject("LDAP://RootDSE").Get("DefaultNamingContext")

If (IsEmpty(strDefaultDomainNC)) Then
  Wscript.Echo("")
  Wscript.Echo("Error: Did not get the Default Naming Context")
  Call Cleanup(2)
End If

'Dispaly Server Name and Default Naming Context
WScript.Echo("")
WScript.Echo("The Computer Name is: " & strServerName)
WScript.Echo("The Domain Naming Context is: " & strDefaultDomainNC)

'Set up the ADO connection required to implement the search.
Set objADOConn = CreateObject("ADODB.Connection")

objADOConn.Provider = "ADsDSOObject"
'Connect using current user credentials
objADOConn.Open "Active Directory Provider"

'Code to demonstrate connecting using alternate credentials
'objADOConn.Open "", _
'  "CN=Administrator,CN=Users," & strDefaultDomainNC, "password"

'Verify successful connection state
If objADOConn.State = adStateOpen Then
  Wscript.Echo("")
  WScript.Echo("Authentication Successful!")
Else
  Wscript.Echo("Authentication Failed.")
  Call Cleanup(3)
End If

Set objADOCommand = CreateObject("ADODB.Command")
Set objADOCommand.ActiveConnection = objADOConn

'Format search criteria using SQL syntax
strADSQuery = "SELECT samAccountName, givenName, sn, AdsPath FROM 'LDAP:// " & _
  strDefaultDomainNC & "' WHERE samAccountName = '" & strSearch & "'"

objADOCommand.CommandText = strADSQuery

'Execute the search
Set objQueryResultSet = objADOCommand.Execute

If (objQueryResultSet.EOF) Then
  Wscript.Echo("")
  WScript.Echo("User " & strSearch & " was not found")
  Call Cleanup(4)
End If

'Gather and echo some general information about the user
WScript.Echo("")
Wscript.Echo "Info for search criteria " & strSearch
WScript.Echo("")

intCount = 0
While Not objQueryResultSet.EOF
  strAdsPath = objQueryResultSet.Fields("AdsPath")
  WScript.Echo("The AdsPath is: " & strAdsPath)
  WScript.Echo("")
  intCount = intCount + 1
  objQueryResultSet.MoveNext
Wend
objADOConn.Close

'This section is for informational purposes
'If a single user was returned, demonstrate binding to the user object
If intCount = 1  Then
  Set objUser = GetObject(strAdsPath)

  'Note: If an attribute is empty, the entire line will not echo
  Wscript.Echo("sAMAccountName: " & objUser.Get("sAMAccountName"))
  WScript.Echo("The Given Name is: " & objQueryResultSet.Fields("givenName"))
  WScript.Echo("The Surname is: " & objQueryResultSet.Fields("sn"))
  WScript.Echo("The AdsPath is: " & objQueryResultSet.Fields("AdsPath"))
  Wscript.Echo("Description:    " & objUser.Get("Description"))
  Wscript.Echo("E-Mail:         " & objUser.Get("mail"))

  Set objUser = Nothing
End If

Dim objShell
SET objShell = CREATEOBJECT("Wscript.Shell")
objShell.Run "cscript.exe \\jc1wsalt03\library\packages\dantools\WriteScriptLog.vbs ""Search Active Directory"""
msgBox "Complete"

Call Cleanup (0)	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Sub ShowUsage()
'
' Purpose: Shows the correct usage to the user.
'
' Input:  None
'
' Output: Help messages are displayed on screen.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowUsage()

Call WScript.Echo(vbCrLf & _
  "Script usage: cscript SearchDC.vbs search_criteria" & vbCrLf & vbCrLf & _
  "Example 1: cscript SearchDC.vbs john" & vbCrLf & _
  "           Will search for and return a user named john." & vbCrLf & _
  "Example 2: cscript SearchDC.vbs j*" & vbCrLf & _
  "           Will search for and return all users with a login" & vbCrLf & _
  "           name begining with j." & vbCrLf & _
  "Example 3: cscript SearchDC.vbs *" & vbCrLf & _
  "           Will search for and return all user logins, groups" & vbCrLf & _
  "           and computer accounts." & vbCrLf & _
  "Example 4: cscript SearchDC.vbs * > results.txt" & vbCrLf & _
  "           Will search for all user logins and computer accounts" & vbCrLf & _
  "           and return the results to a file named results.txt.")

Call Cleanup(0)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Sub Cleanup
'
' Purpose: Cleanup objects and exit
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Cleanup(intExitCode)

  Set objADOConn = Nothing
  Wscript.Quit(intExitCode)

End Sub