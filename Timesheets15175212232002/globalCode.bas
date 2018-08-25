Attribute VB_Name = "globalCode"
Option Explicit

Public DETimesheets As New ADODB.Connection
Public Const strConnectionStart = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const strConnectionEnd = ";Mode=Read|Write;Persist Security Info=False"
Public Const txtIniFile = "timesheets.ini"
Public strConnection As String


'
'API Declares
'
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public intFormAction As Integer
Public Const ADD_NEW = 1
Public Const EDIT = 2
Public Const DELETE = 3

Public Const NORMAL_USER = 1
Public Const PROJECT_MANAGER = 2
Public Const FINANCIAL_USER = 3
Public Const SUPER_USER = 4

Public Const strCaption = "Timesheets Lite"

Public cmdSelectProjects As New ADODB.Command
Public cmdSelectProjectsOrderByID As New ADODB.Command
Public cmdSelectUsers As New ADODB.Command
Public cmdSelectProjectByID As New ADODB.Command
Public cmdSelectUserByID As New ADODB.Command
Public cmdSelectUserTimesDate As New ADODB.Command
Public cmdSelectTimeByID As New ADODB.Command
Public cmdSelectOutstandingProjectTimes As New ADODB.Command
Public cmdTimeReport As New ADODB.Command
Public cmdSelectUserProjectTimes As New ADODB.Command
Public cmdSelectTimesByUser As New ADODB.Command
Public cmdSelectTimesByProject As New ADODB.Command
Public cmdSelectLoggedInUsers As New ADODB.Command
Public cmdSelectLoggedInUserByID As New ADODB.Command

Public proCurrent As clsProject
Public proOld As clsProject
Public usrCurrent As clsUser
Public usrOld As clsUser
Public usrLoggedIn As clsUser
Public timCurrent As clsTime
Public timOld As clsTime

Sub Main()
  initialise
End Sub

Sub initialise()
  Dim lngTickStart As Long
  lngTickStart = GetTickCount
  frmSplash.Show
  frmSplash.Refresh
  While (GetTickCount - lngTickStart) < 1000
  Wend
  makeConnection
  defineCommandObjects
  
  Set proCurrent = New clsProject
  Set proOld = New clsProject

  Set usrCurrent = New clsUser
  Set usrOld = New clsUser
  
  Set usrLoggedIn = New clsUser

  Set timCurrent = New clsTime
  Set timOld = New clsTime
  Unload frmSplash
  frmMDIMain.Show
End Sub

Private Sub makeConnection()
  Dim strReadIni As String
  
  'INI File Processing
  '
  '
  'Read INI file and get database location
  '
  strReadIni = readINIFile("DatabaseLoc")
  '
  'Construct database connection string on the basis of this, note that if no string is returned then
  'assume a default location
  '
  If Len(strReadIni) = 0 Then
    strConnection = strConnectionStart & App.Path & "\" & "timesheets.mdb" & strConnectionEnd
  Else
    strConnection = strConnectionStart & strReadIni & strConnectionEnd
  End If
  DETimesheets.ConnectionString = strConnection
  DETimesheets.CursorLocation = adUseClient
  DETimesheets.Open strConnection, "Admin"
End Sub

Public Sub exitSystem()
  '
  'Close off the database connections
  '
  If usrLoggedIn.lngUserID > 0 Then securityCode.removeLoggedInUser usrLoggedIn.lngUserID
  DETimesheets.Close
  Set DETimesheets = Nothing
  '
  'Release Form Memory
  '
  Set frmMDIMain = Nothing
  Set frmProjectMaintenance = Nothing
  Set frmUserMaintenance = Nothing
  Set frmTimeSheet = Nothing
  End

End Sub

Public Sub defineCommandObjects()
  Dim strSQL As String
  
  With cmdSelectProjects
    .ActiveConnection = DETimesheets
    .CommandText = "SELECT * from `Project Details` ORDER BY `Project Details`.datCreated"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectProjectsOrderByID
    .ActiveConnection = DETimesheets
    .CommandText = "SELECT * from `Project Details` ORDER BY `Project Details`.lngProjectID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With

  With cmdSelectUsers
    .ActiveConnection = DETimesheets
    .CommandText = "SELECT * from `User Details` ORDER BY `User Details`.strLastName"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectProjectByID
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS ID Text; SELECT * from `Project Details` where `Project Details`.lngProjectID =ID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectUserByID
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS ID Text; SELECT * from `User Details` where `User Details`.lngUserID=ID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
    
  With cmdSelectTimeByID
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS ID Text; SELECT * from " & _
    "Times where Times.lngID = ID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With

  
  With cmdSelectOutstandingProjectTimes
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS StartDate DateTime,EndDate DateTime; SELECT * from Times " & _
    "WHERE ((Times.datDate BETWEEN StartDate AND EndDate) AND Times.blnInvoiced=False);"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
    
  With cmdTimeReport
    .ActiveConnection = DETimesheets
    .CommandText = "Parameters StartDate DateTime,EndDate DateTime; " & _
     "SELECT [Project Details].strCustomer, Times.lngProjectID, [Project Details].strProjectName, " & _
     "[Project Details].strProjectNumber, Times.lngUserID, [User Details].strFirstName, [User Details].strLastName, " & _
     "Sum(Times.dblHours) AS SumOfdblHours " & _
     "FROM [User Details] INNER JOIN ([Project Details] INNER JOIN Times ON [Project Details].lngProjectID = Times.lngProjectID) ON [User Details].lngUserID = Times.lngUserID " & _
     "WHERE ((Times.datDate Between StartDate And EndDate) AND (Times.blnInvoiced=False)) " & _
     "GROUP BY [Project Details].strCustomer, Times.lngProjectID, [Project Details].strProjectName, [Project Details].strProjectNumber, Times.lngUserID, [User Details].strFirstName, [User Details].strLastName;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectUserProjectTimes
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS ProjectID Text,UserID Text; SELECT * from Times " & _
    "WHERE (Times.lngProjectID = ProjectID AND Times.lngUserID=UserID) " & _
    "ORDER BY Times.datDate;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
 
  With cmdSelectUserTimesDate
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS UserID Text,SelectDate DateTime; SELECT * from Times " & _
    "WHERE (Times.lngUserID=UserID AND Times.datDate=SelectDate) " & _
    "ORDER BY Times.datDate;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
   
  With cmdSelectTimesByUser
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS UserID Text; SELECT * from Times " & _
    "WHERE Times.lngUserID=UserID;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectTimesByProject
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS ProjectID Text; SELECT * from Times " & _
    "WHERE Times.lngProjectID=ProjectID;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With

  With cmdSelectLoggedInUsers
    .ActiveConnection = DETimesheets
    .CommandText = "SELECT * from [Users Logged In]"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectLoggedInUserByID
    .ActiveConnection = DETimesheets
    .CommandText = "PARAMETERS UserID Text; SELECT * from [Users Logged In] " & _
    "WHERE [Users Logged In].lngUserID=UserID;"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
End Sub

Function returnRS(cmdCommand As ADODB.Command) As ADODB.Recordset
  Dim rstReturnRS As New ADODB.Recordset
  
  With rstReturnRS
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmdCommand
  End With
  Set returnRS = rstReturnRS
End Function

Public Function fillProjectListView(lvwBox As ListView, Optional lngSelectedItem As Long)
  Dim rstProjects As New ADODB.Recordset
  Dim intListIndex As Integer
  Dim lstItem As ListItem
  
  Set rstProjects = returnRS(cmdSelectProjects)
  
  lvwBox.ListItems.clear
  If rstProjects.EOF = False Then
    rstProjects.MoveFirst
    While rstProjects.EOF <> True
      Set lstItem = lvwBox.ListItems.Add
      lstItem = rstProjects![strProjectNumber]
      lstItem.SubItems(1) = rstProjects![strProjectName]
      lstItem.Key = "Item_" & rstProjects![lngProjectID]
      If rstProjects![lngProjectID] = lngSelectedItem Then intListIndex = lstItem.Index
      rstProjects.MoveNext
    Wend
  End If
  rstProjects.Close
  Set rstProjects = Nothing
  If lngSelectedItem > 0 Then
    Set lvwBox.SelectedItem = lvwBox.ListItems(intListIndex)
  End If
End Function

Public Function fillUserListView(lvwBox As ListView)
  Dim rstUsers As New ADODB.Recordset
  Dim lstItem As ListItem
  
  Set rstUsers = returnRS(cmdSelectUsers)
  
  lvwBox.ListItems.clear
  If rstUsers.EOF = False Then
    rstUsers.MoveFirst
    While rstUsers.EOF <> True
      Set lstItem = lvwBox.ListItems.Add
      lstItem = rstUsers![strFirstName] & " " & rstUsers![strLastName]
      lstItem.SubItems(1) = rstUsers![strAbbreviation]
      lstItem.Key = "Item_" & rstUsers![lngUserID]
      rstUsers.MoveNext
    Wend
  End If
  rstUsers.Close
  Set rstUsers = Nothing
End Function

Public Function checkSelected(lvwTmp As Variant) As Integer
  Dim intLoop As Integer
  checkSelected = -1
  If lvwTmp.ListItems.Count = 0 Then
    checkSelected = -1
    Exit Function
  End If
  For intLoop = 1 To lvwTmp.ListItems.Count
    If lvwTmp.ListItems(intLoop).Selected = True Then
      checkSelected = intLoop
      Exit For
    End If
  Next intLoop
End Function

Public Function parseDirectory(strFilename) As String
  Dim strTmp As String
  Dim lngPosition As Long
  
  If IsNull(strFilename) = False Then
    strTmp = StrReverse(strFilename)
    lngPosition = InStr(strTmp, "\")
    strTmp = Left(strFilename, Len(strFilename) - lngPosition)
    parseDirectory = strTmp
    If Len(parseDirectory) = 0 Then
      parseDirectory = App.Path
    End If
  Else
    parseDirectory = App.Path
  End If
End Function

Public Function getIDFromKey(strKey As String) As Long
  getIDFromKey = CLng(Right(strKey, Len(strKey) - Len("Item_")))
End Function
