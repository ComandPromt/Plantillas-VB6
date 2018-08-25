Attribute VB_Name = "timesheetCode"
Option Explicit

Public intTimesheetWeekday As Integer
Public datSelectedDate As Date


Public Sub deleteProjectTimes(lngProjectID As Long)
  Dim rstTimes As New ADODB.Recordset
  Dim intCount As Integer
  
  cmdSelectTimesByProject.Parameters(0) = lngProjectID
  Set rstTimes = returnRS(cmdSelectTimesByProject)
  If rstTimes.EOF <> True Then
    rstTimes.MoveFirst
    While rstTimes.EOF <> True
      intCount = intCount + 1
      rstTimes.DELETE
      rstTimes.MoveNext
    Wend
    rstTimes.UpdateBatch
  End If
  rstTimes.Close
  Set rstTimes = Nothing
  Debug.Print "#timesheetCode#deleteProjectTimes " & intCount & " records deleted"
End Sub


Public Sub deleteUserTimes(lngUserID As Long)
  Dim rstTimes As New ADODB.Recordset
  Dim intCount As Integer
  
  cmdSelectTimesByUser.Parameters(0) = lngUserID
  Set rstTimes = returnRS(cmdSelectTimesByUser)
  If rstTimes.EOF <> True Then
    rstTimes.MoveFirst
    While rstTimes.EOF <> True
      intCount = intCount + 1
      rstTimes.DELETE
      rstTimes.MoveNext
    Wend
    rstTimes.UpdateBatch
  End If
  rstTimes.Close
  Set rstTimes = Nothing
  Debug.Print "#timesheetCode#deleteUserTimes " & intCount & " records deleted"
End Sub

Public Function fillDateCombo(cboTmp As ComboBox)
  Dim intWeeks As Integer
  Dim datLastMonday As Date
  cboTmp.clear
  
  datLastMonday = Date - Weekday(Date, vbMonday) + 1
  cboTmp.AddItem "Monday " & Day(datLastMonday) & "/" & Month(datLastMonday) & "/" & Year(datLastMonday)
  For intWeeks = 1 To 8
    cboTmp.AddItem "Monday " & Day(datLastMonday - intWeeks * 7) & "/" & Month(datLastMonday - intWeeks * 7) & "/" & Year(datLastMonday - intWeeks * 7)
  Next intWeeks
End Function

Public Function fillTimesListview(lvwTimes As ListView, cboTmp As ComboBox)
  
  Dim rstTimes As New ADODB.Recordset
  Dim itmTmp As ListItem
  Dim strTmp As String
  intFormAction = 0
  lvwTimes.ListItems.clear
  If cboTmp.ListIndex = -1 Then
    strTmp = Right(cboTmp.List(0), Len(cboTmp.List(0)) - Len("Monday") - 1)
    datSelectedDate = returnDate(CDate(strTmp))
  Else
    strTmp = Right(cboTmp.List(cboTmp.ListIndex), Len(cboTmp.List(cboTmp.ListIndex)) - Len("Monday") - 1)
    datSelectedDate = returnDate(CDate(strTmp))
  End If
  cmdSelectUserTimesDate.Parameters(0) = usrLoggedIn.lngUserID
  cmdSelectUserTimesDate.Parameters(1) = datSelectedDate
  Set rstTimes = returnRS(cmdSelectUserTimesDate)
  If rstTimes.EOF <> True Then
    rstTimes.MoveFirst
    While rstTimes.EOF <> True
      Set itmTmp = lvwTimes.ListItems.Add
      itmTmp = returnProjectNumberWithID(rstTimes![lngProjectID])
      itmTmp.Key = "Item_" & rstTimes![lngID]
      itmTmp.SubItems(1) = rstTimes![memDescription]
      itmTmp.SubItems(2) = rstTimes![dblHours]
      rstTimes.MoveNext
    Wend
  End If
  updateTotalHours lvwTimes
  rstTimes.Close
  Set rstTimes = Nothing
End Function

Public Function fillProjectCombo(cboTmp As ComboBox)
  Dim rstProjects As New ADODB.Recordset
  Dim intItem As Integer
  cboTmp.clear
  Set rstProjects = returnRS(cmdSelectProjects)
  If rstProjects.EOF <> True Then
    rstProjects.MoveFirst
    intItem = 0
    While rstProjects.EOF <> True
      cboTmp.AddItem rstProjects![strProjectNumber] & " " & rstProjects![strProjectName]
      cboTmp.ItemData(intItem) = rstProjects![lngProjectID]
      rstProjects.MoveNext
      intItem = intItem + 1
    Wend
  End If
  rstProjects.Close
  Set rstProjects = Nothing
End Function

Public Function returnProjectNumberWithID(lngProjectID As Long) As String
  Dim rstProject As New ADODB.Recordset
  
  cmdSelectProjectByID.Parameters(0) = lngProjectID
  Set rstProject = returnRS(cmdSelectProjectByID)
  If rstProject.EOF <> True Then
    returnProjectNumberWithID = rstProject![strProjectNumber]
  Else
    returnProjectNumberWithID = ""
  End If
  rstProject.Close
  Set rstProject = Nothing
End Function

Public Function returnProjectNameWithID(lngProjectID As Long) As String
  Dim rstProject As New ADODB.Recordset
  
  cmdSelectProjectByID.Parameters(0) = lngProjectID
  Set rstProject = returnRS(cmdSelectProjectByID)
  If rstProject.EOF <> True Then
    returnProjectNameWithID = rstProject![strProjectName]
  Else
    returnProjectNameWithID = ""
  End If
  rstProject.Close
  Set rstProject = Nothing
End Function

Private Function returnDate(datWeekStart As Date) As Date
  Select Case intTimesheetWeekday
    Case vbMonday
      returnDate = datWeekStart
    Case vbTuesday
      returnDate = datWeekStart + 1
    Case vbWednesday
      returnDate = datWeekStart + 2
   Case vbThursday
      returnDate = datWeekStart + 3
    Case vbFriday
      returnDate = datWeekStart + 4
    Case vbSaturday
      returnDate = datWeekStart + 5
    Case vbSunday
      returnDate = datWeekStart + 6
  End Select
End Function

Public Function enableDisplay()
  frmTimeSheet.cboWeek.Enabled = True
  frmTimeSheet.fraDay.Enabled = True
  frmTimeSheet.lvwTimes.Enabled = True
  frmTimeSheet.cboProject.Enabled = True
  frmTimeSheet.txtDescription.Enabled = True
  frmTimeSheet.txtTime.Enabled = True
  frmTimeSheet.btnAddEntry.Enabled = True
  frmTimeSheet.btnDeleteEntry.Enabled = True
End Function

Public Function disableDisplay()
  frmTimeSheet.cboWeek.Enabled = False
  frmTimeSheet.fraDay.Enabled = False
  frmTimeSheet.lvwTimes.Enabled = False
  frmTimeSheet.cboProject.Enabled = False
  frmTimeSheet.txtDescription.Enabled = False
  frmTimeSheet.txtTime.Enabled = False
  frmTimeSheet.btnAddEntry.Enabled = False
  frmTimeSheet.btnDeleteEntry.Enabled = False
End Function

Public Function enableTimeDisplay()
  frmTimeSheet.cboProject.ListIndex = -1
  frmTimeSheet.txtDescription = ""
  frmTimeSheet.txtTime = 0
  frmTimeSheet.cboProject.Enabled = True
  frmTimeSheet.txtDescription.Enabled = True
  frmTimeSheet.txtTime.Enabled = True
End Function

Public Function disableTimeDisplay()
  frmTimeSheet.cboProject.ListIndex = -1
  frmTimeSheet.txtDescription = ""
  frmTimeSheet.txtTime = 0
  frmTimeSheet.cboProject.Enabled = False
  frmTimeSheet.txtDescription.Enabled = False
  frmTimeSheet.txtTime.Enabled = False
End Function

Public Function returnProjectID(cboTmp As ComboBox) As Long
  If cboTmp.ListIndex <> -1 Then
    returnProjectID = cboTmp.ItemData(cboTmp.ListIndex)
  Else
    returnProjectID = 0
  End If
End Function

Public Function selectCorrectProject(lngProjectID As Long, cboTmp As ComboBox)
  Dim intLoop As Integer
  For intLoop = 0 To cboTmp.ListCount - 1
    If lngProjectID = cboTmp.ItemData(intLoop) Then
      cboTmp.ListIndex = intLoop
      Exit Function
    End If
  Next intLoop
End Function

Private Function updateTotalHours(lvwTmp As ListView)
  Dim intLoop As Integer
  Dim dblRunningTotal As Double
  dblRunningTotal = 0
  For intLoop = 1 To lvwTmp.ListItems.Count
    dblRunningTotal = dblRunningTotal + lvwTmp.ListItems(intLoop).SubItems(2)
  Next intLoop
  frmTimeSheet.lblTotal = dblRunningTotal & " hrs"
End Function
