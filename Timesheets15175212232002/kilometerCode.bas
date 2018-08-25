Attribute VB_Name = "kilometerCode"
Option Explicit

Public Function fillKilometerList(lngProjectID As Long, lvwKilometers As ListView)
  Dim rstKilometers As New ADODB.Recordset
  Dim itmTmp As ListItem
  Dim strTmp As String
  intFormAction = 0
  lvwKilometers.ListItems.clear
  
  cmdSelectUserProjectKilometers.Parameters(0) = lngProjectID
  cmdSelectUserProjectKilometers.Parameters(1) = usrLoggedIn.lngUserID
  Set rstKilometers = returnRS(cmdSelectUserProjectKilometers)
  If rstKilometers.EOF <> True Then
    rstKilometers.MoveFirst
    While rstKilometers.EOF <> True
      Set itmTmp = lvwKilometers.ListItems.Add
      itmTmp = rstKilometers![lngID]
      itmTmp.SubItems(1) = rstKilometers![datDate]
      itmTmp.SubItems(2) = rstKilometers![memDescription]
      itmTmp.SubItems(3) = rstKilometers![dblKilometers]
      rstKilometers.MoveNext
    Wend
  End If
  updateTotalKilometers lvwKilometers
  rstKilometers.Close
  Set rstKilometers = Nothing
End Function

Public Function enableDisplay()
  frmKilometers.txtDate.Enabled = True
  frmKilometers.txtKilometers.Enabled = True
  frmKilometers.txtDescription.Enabled = True
  frmKilometers.txtDate = 0
  frmKilometers.txtKilometers = 0
  frmKilometers.txtDescription = ""
End Function

Public Function disableDisplay()
  frmKilometers.txtDate = 0
  frmKilometers.txtKilometers = 0
  frmKilometers.txtDescription = ""
  frmKilometers.txtDate.Enabled = False
  frmKilometers.txtKilometers.Enabled = False
  frmKilometers.txtDescription.Enabled = False
End Function

Private Function updateTotalKilometers(lvwTmp As ListView)
  Dim intLoop As Integer
  Dim dblRunningTotal As Double
  dblRunningTotal = 0
  For intLoop = 1 To lvwTmp.ListItems.Count
    dblRunningTotal = dblRunningTotal + lvwTmp.ListItems(intLoop).SubItems(3)
  Next intLoop
  frmKilometers.lblTotalKilometers = dblRunningTotal & " kms"

End Function
