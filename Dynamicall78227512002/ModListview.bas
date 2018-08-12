Attribute VB_Name = "ModListview"
'Designed and developed by Chris Hatton. if you want to reuse this code please email me.
'chris@hatton.com


Public Sub PopluateListView(SQL As String, Lv As ListView)
On Error Resume Next
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Long
Dim j As Long
Dim k As Long
rs.Open SQL, cn, adOpenStatic, adLockOptimistic
FrmView.Caption = "Viewing Table " & FrmMain.TableView1.SelectedItem.Text
    For i = 0 To rs.Fields.Count - 1
        Lv.ColumnHeaders.Add , , rs.Fields(i).Name
       
     
    Next i
   FrmMain.ProgressBar1.Max = rs.RecordCount 'busy signals
   FrmMain.ProgressBar1.Value = 0
   Screen.MousePointer = 11
    For k = 0 To rs.RecordCount
    
    
        If IsNull(rs.Fields(0).Value) Then _
         Lv.ListItems.Add , , "" Else Lv.ListItems.Add , , rs.Fields(0).Value
         
         
         For j = 1 To rs.Fields.Count
            
            If IsNull(rs.Fields(1).Value) Then _
            Lv.ListItems.Item(k + 1).ListSubItems.Add , , "" Else Lv.ListItems.Item(k + 1).ListSubItems.Add , , rs.Fields(j).Value
         
         
         Next j
         
         FrmMain.ProgressBar1.Value = k
         rs.MoveNext
 
 Next k
FrmView.StatusBar1.Panels(1).Text = "Record(s) " & rs.RecordCount
FrmMain.ProgressBar1.Value = 0 'not so busy
Screen.MousePointer = 0
rs.Close
Set rs = Nothing





End Sub
