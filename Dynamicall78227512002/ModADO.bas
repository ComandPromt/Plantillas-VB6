Attribute VB_Name = "ModADO"
'Designed and developed by Chris Hatton. if you want to reuse this code please email me.
'chris@hatton.com

Public TvGrp(0 To 100) As Node
Private FirstLoad As Boolean 'first time loading this app
Public Sub ScopeRecords(SQL As String, DBTV As TreeView)

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Dim i As Long
Dim j As Long
Screen.MousePointer = 11
On Error GoTo PopulateError

rs.Open SQL, cn, adOpenStatic, adLockOptimistic
With FrmMain
.ProgressBar1.Value = 0
.ProgressBar1.Max = rs.RecordCount
DBTV.Nodes.Clear
For i = 0 To rs.Fields.Count - 1
            DoEvents
            .Label2 = "Loading Field " & i + 1 & " of " & rs.Fields.Count
            DoEvents
            
      
    Set TvGrp(i) = DBTV.Nodes.Add(, , , rs.Fields(i).Name, 1)
    
        If IsNull(rs.Fields(i).Value) Then _
    Set TvGrp(i) = DBTV.Nodes.Add(TvGrp(i), tvwChild, , "", 4) Else _
    Set TvGrp(i) = DBTV.Nodes.Add(TvGrp(i), tvwChild, , rs.Fields(i).Value, 4)
   
   
        For j = 2 To rs.RecordCount
        rs.MoveNext
            If Not Len(rs.Fields(i).Value) = 0 Then _
            Set TvGrp(i) = DBTV.Nodes.Add(TvGrp(i), tvwNext, , rs.Fields(i).Value, 4)
                .ProgressBar1.Value = j - 1
        Next j
    
    rs.MoveFirst

.ProgressBar1.Value = 0

Next i


rs.Close
Set rs = Nothing

Screen.MousePointer = 0
.Label2 = ""

End With
Exit Sub

PopulateError:
Screen.MousePointer = 0


MsgBox "Found Errors Trying to Popluate Control" & vbNewLine & Err.Description, vbCritical + vbOKOnly


End Sub
Public Sub AddLocalValue(SQL, Value As String, TableName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

rs.AddNew

rs.Fields(TableName).Value = Value
rs.Update
rs.Close
Set rs = Nothing

End Sub

Public Sub AddLocalSubValue(SQL, SubValue As String, RsRow As Long, TableName As String, SubTv As TreeView)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

If SubValue = "" Then GoTo CancelUpdate
    
    rs.Move RsRow, 1
    rs.Fields(TableName).Value = SubValue
    rs.Update

    
    SubTv.Nodes.Add SubTv.SelectedItem.Parent.Index, tvwChild, , SubValue, 4
    SubTv.Refresh


    
CancelUpdate:
rs.Close
Set rs = Nothing
End Sub
Public Sub DelLocalValue(SQL, RsRow As Long, TableName As String)
On Error GoTo DelUpdateErr
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

rs.Move RsRow, 1
rs.Delete
rs.Update
rs.Close
Set rs = Nothing
Exit Sub
DelUpdateErr:
MsgBox "Could not Delete", vbCritical + vbOKOnly

End Sub

Public Sub EditLocalValue(SQL, Value As String, RsRow As Long, TableName As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.Open SQL, cn, adOpenKeyset, adLockOptimistic

rs.Move RsRow, 1
rs.Fields(TableName).Value = Value
rs.Update
rs.Close
Set rs = Nothing
End Sub
Sub ClearMenuItems()
Dim j As Long
    
    For j = 1 To FrmMain.dbitem.Count - 1
 
        Unload FrmMain.dbitem(j)
        
    Next j
  FrmMain.dbitem(0).Caption = ""
End Sub

Public Sub GetLocalTables(LvTable As ListView)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
Dim i As Long
     
     LvTable.ListItems.Clear
     LvTable.ColumnHeaders.Clear
     LvTable.ColumnHeaders.Add , , "Tables", 2600
     LvTable.ColumnHeaders.Add , , "Total Records", 1500
     LvTable.ColumnHeaders.Add , , "Total Fields", 1500
     LvTable.View = lvwReport
    
            For i = 0 To rs.Fields.Count
                LvTable.ListItems.Add , , rs!Table_Name, , 2
                    
                        If Not i = 0 Then Load FrmMain.dbitem.Item(i)
                        FrmMain.dbitem.Item(i).Caption = rs!Table_Name
                    
                    rs.MoveNext

            Next i
                

rs.Close
Set rs = Nothing

End Sub

Public Sub GetTableSatisitic(LvTable As ListView)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer


With FrmMain
    For i = 1 To LvTable.ListItems.Count
    rs.Open "select * from [" & LvTable.ListItems.Item(i).Text & "]", cn, adOpenStatic, adLockOptimistic

        LvTable.ListItems.Item(i).ListSubItems.Add , , rs.RecordCount
        LvTable.ListItems.Item(i).ListSubItems.Add , , rs.Fields.Count
    rs.Close
    Next i
End With


Set rs = Nothing





End Sub
